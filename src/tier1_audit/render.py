"""Markdown + JSON + HTML renderers.

All deterministic: same AuditReport in -> byte-identical bytes out.

Polish round (Stage 2 task 04):
- Restructured section ordering: cover, exec summary, TOC, then content.
  Pillars and magic-number anomalies are promoted from `## 6.5` / `## 6.6`
  into top-level H2 sections (since they are the headline findings).
- Three Mermaid diagrams embedded:
    1. Sheet data flow (graph LR with cross-sheet reference counts)
    2. VBA classification overview (graph TB grouping modules under labels)
    3. Pillar cell impact (graph LR connecting top pillars to affected sheets)
- HTML render with embedded Mermaid (CDN by default, --mermaid-inline downloads
  and inlines the script for fully offline viewing).
- Domain-detector hint shown in the Executive Summary block.
"""

from __future__ import annotations

import html as _html
import json
import re
from collections import defaultdict
from dataclasses import asdict
from typing import Any

from .smells import SMELL_TYPES, TRIVIAL_NUMBERS

# Field names to omit from JSON serialization (huge or noisy or internal-only)
_EXCLUDED_FIELDS = frozenset({"source_text", "_sheet_edges"})


# ---------------------------------------------------------------------------
# JSON
# ---------------------------------------------------------------------------

def _to_jsonable(obj: Any) -> Any:
    if hasattr(obj, "__dataclass_fields__"):
        return {
            k: _to_jsonable(v)
            for k, v in asdict(obj).items()
            if k not in _EXCLUDED_FIELDS
        }
    if isinstance(obj, dict):
        return {
            k: _to_jsonable(v)
            for k, v in obj.items()
            if k not in _EXCLUDED_FIELDS
        }
    if isinstance(obj, (list, tuple)):
        return [_to_jsonable(x) for x in obj]
    if isinstance(obj, frozenset):
        return sorted(obj)
    return obj


def render_json(report) -> str:
    data = _to_jsonable(report)
    return json.dumps(data, sort_keys=True, indent=2, ensure_ascii=False) + "\n"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _md_escape(s: str) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


_RE_NON_ASCII_ID = re.compile(r"[^A-Za-z0-9_]")


def _safe_node_id(s: str, prefix: str = "n") -> str:
    """Mermaid node IDs must be ASCII-clean. Build a deterministic slug.

    We map non-ASCII to underscores and prefix to avoid collisions with
    pure-numeric IDs.
    """
    slug = _RE_NON_ASCII_ID.sub("_", s)
    if not slug:
        slug = "x"
    # Ensure it starts with a letter to satisfy Mermaid
    if not slug[0].isalpha():
        slug = prefix + slug
    return slug


def _safe_node_label(s: str) -> str:
    """Sanitize a label that goes inside Mermaid `["..."]` brackets.

    The label is rendered as text by Mermaid; the only characters that
    actually break parsing inside a quoted label are the closing quote
    and the structural `]` that terminates the bracket. Replace those.
    Other characters (parentheses, CJK, punctuation) render fine on
    Mermaid 10.x with securityLevel:'loose'.
    """
    if not s:
        return s
    return (str(s)
            .replace('"', "'")
            .replace("]", ")"))


def _slugify_anchor(s: str) -> str:
    """Build a markdown anchor slug from a header. GFM-style: lowercase,
    spaces -> hyphens, drop punctuation. Used for TOC links."""
    s = s.lower()
    s = re.sub(r"[^a-z0-9\s\-]", "", s)
    s = re.sub(r"\s+", "-", s.strip())
    return s


def _sub_score_bar(value: int, width: int = 10) -> str:
    """Tiny ASCII sparkline-like bar for a 0-20 sub-score."""
    filled = int(round(value / 20 * width))
    return "█" * filled + "░" * (width - filled)


# ---------------------------------------------------------------------------
# Headline-finding extraction (used by both md and html exec summary)
# ---------------------------------------------------------------------------

def _headline_findings(report) -> list:
    """Return up to 3 single-line headline findings: Top pillar, top anomaly
    or top smell, plus one risk indicator. Strings only — no markup beyond
    inline backticks."""
    out: list = []
    if report.pillars:
        p = report.pillars[0]
        if p.member_count > 1:
            out.append(
                f"**Pillar group**: `{p.location}` ({p.member_count} cells) "
                f"each cascade into {p.fan_in} formulas."
            )
        else:
            out.append(
                f"**Pillar**: `{p.location}` cascades into {p.fan_in} formulas across "
                f"{p.affected_sheet_count} sheet(s)."
            )

    if report.anomalies:
        a = report.anomalies[0]
        out.append(
            f"**Anomaly**: position #{a.position_index + 1} of a {a.cluster_size}-cell cluster "
            f"normally `{a.mode_value}` but `{a.outlier_value}` at "
            f"`{a.outlier_locations[0]}` ({a.confidence} conf)."
        )
    elif report.smells:
        # Top smell by metric
        top_smell = max(report.smells, key=lambda s: s.metric)
        out.append(
            f"**Top smell**: `{top_smell.smell_type}` at `{top_smell.location}` "
            f"(metric={top_smell.metric:g}, severity={top_smell.severity})."
        )

    r = report.risk_indicators
    risks = []
    if r.very_hidden_sheets:
        risks.append(f"{len(r.very_hidden_sheets)} veryHidden sheet(s)")
    if r.hidden_sheets:
        risks.append(f"{len(r.hidden_sheets)} hidden sheet(s)")
    if r.cells_with_errors:
        risks.append(f"{len(r.cells_with_errors)} formula error cell(s)")
    if r.external_workbook_references:
        risks.append(f"{len(r.external_workbook_references)} external workbook ref(s)")
    if r.circular_reference_suspects:
        risks.append(f"{len(r.circular_reference_suspects)} circular suspect(s)")
    if risks:
        out.append("**Risks**: " + "; ".join(risks) + ".")

    return out[:3]


# ---------------------------------------------------------------------------
# Mermaid diagrams — three of them
# ---------------------------------------------------------------------------

# Caps to keep diagrams readable and ASCII-clean.
# D6: hard cap at 15 nodes per diagram (any single diagram).
_DIAG1_MAX_NODES = 15
_DIAG1_MAX_EDGES = 25
_DIAG3_MAX_PILLARS = 5


def _build_sheet_dataflow_mermaid(report) -> str:
    """Diagram 1 — graph LR of cross-sheet formula references.

    Nodes are sheets, color-coded by visibility. Edges have count labels.
    Top-N most-connected sheets only.
    """
    # Build a sheet -> sheet edge weight from the cell_rows / formulas:
    # We don't have raw cell rows here, so we rely on the smells.detect_multiple_references
    # outgoing/incoming map indirectly — but those are sheet|ref keyed.
    # Instead, walk smells of type 'multiple-references' (which carry source/target
    # sheet info via location) — but smells already collapse to target-only.
    # Solution: re-derive from `report.sheets` + named ranges + smells; we use the
    # cross_sheet_reference_count summary plus a rough per-sheet attribution by
    # walking smells whose evidence mentions cross-sheet via the
    # `incoming` reverse map. Since we don't have that here either, we use a
    # simpler heuristic: build edges by looking at the duplicated-formulas
    # smell `evidence` strings (which contain sample formulas with sheet
    # references) and the magic-numbers `sample_context`. This is approximate
    # but fully ASCII and deterministic.
    #
    # NOTE: A cleaner approach would be to plumb the raw outgoing-edges map
    # through to render. To stay minimally invasive, we attach it to the
    # report as `report._sheet_edges` in audit.py (see that file for the
    # populated attribute).
    edges = getattr(report, "_sheet_edges", None) or {}
    sheet_state = {s.name: s.state for s in report.sheets}

    # Tally per-node degree to pick top-N
    degree: dict = defaultdict(int)
    for (src, tgt), cnt in edges.items():
        degree[src] += cnt
        degree[tgt] += cnt
    # Always include all sheets, but cap total nodes at _DIAG1_MAX_NODES.
    # When we'd overflow, reserve 1 slot for the "+ N more" placeholder.
    if len(sheet_state) <= _DIAG1_MAX_NODES:
        kept_nodes = sorted(sheet_state.keys())
    else:
        budget = _DIAG1_MAX_NODES - 1  # leave room for "+ N more"
        kept_nodes = [
            n for n, _ in sorted(degree.items(), key=lambda kv: (-kv[1], kv[0]))
        ][:budget]
        # Fold isolated sheets into "+ N more" placeholder later
        kept_nodes = sorted(set(kept_nodes))

    kept_set = set(kept_nodes)
    # Filter edges
    kept_edges_items = [
        (s, t, c) for (s, t), c in edges.items()
        if s in kept_set and t in kept_set and s != t
    ]
    kept_edges_items.sort(key=lambda x: (-x[2], x[0], x[1]))
    if len(kept_edges_items) > _DIAG1_MAX_EDGES:
        kept_edges_items = kept_edges_items[:_DIAG1_MAX_EDGES]
        truncation_note = f"%% truncated to top {_DIAG1_MAX_EDGES} edges"
    else:
        truncation_note = ""

    lines: list = ["graph LR"]
    if truncation_note:
        lines.append(truncation_note)

    # Node IDs and labels (deterministic ordering)
    # Map sheet name -> safe id once
    nid: dict = {}
    used: set = set()
    for n in sorted(kept_nodes):
        base = _safe_node_id(n, prefix="S")
        # Ensure uniqueness (two CJK sheets may collapse to same slug)
        candidate = base
        i = 1
        while candidate in used:
            candidate = f"{base}_{i}"
            i += 1
        nid[n] = candidate
        used.add(candidate)
        label = _safe_node_label(n)
        lines.append(f'{candidate}["{label}"]')

    # Edges
    for src, tgt, cnt in kept_edges_items:
        sa = nid.get(src)
        ta = nid.get(tgt)
        if not sa or not ta:
            continue
        lines.append(f"{sa} -->|{cnt}| {ta}")

    # Color classes by sheet state
    visible_ids = [nid[n] for n in kept_nodes if sheet_state.get(n) == "visible"]
    hidden_ids = [nid[n] for n in kept_nodes if sheet_state.get(n) == "hidden"]
    very_hidden_ids = [nid[n] for n in kept_nodes if sheet_state.get(n) == "veryHidden"]

    if hidden_ids:
        lines.append(
            f"classDef hidden fill:#ffe082,stroke:#b07b00,color:#000000;"
        )
        lines.append(f"class {','.join(hidden_ids)} hidden")
    if very_hidden_ids:
        lines.append(
            f"classDef veryHidden fill:#ef9a9a,stroke:#b71c1c,color:#000000;"
        )
        lines.append(f"class {','.join(very_hidden_ids)} veryHidden")

    # If we omitted nodes, add a tiny "+N more" floating note
    omitted = len(sheet_state) - len(kept_nodes)
    if omitted > 0:
        lines.append(f'more["+ {omitted} more sheets"]:::omitted')
        lines.append(
            "classDef omitted fill:#eeeeee,stroke:#999999,color:#666666;"
        )

    return "\n".join(lines)


_VBA_DIAGRAM_MAX_NODES = 15  # readability cap per Mermaid block (D6)


def _build_vba_classification_mermaid(report) -> str:
    """Diagram 2 — readable graph TB.

    D6 fix: limit to <= 15 nodes per diagram. Show top-2 classifications by
    module count; remaining classifications appear in the per-classification
    summary table preceding this diagram. Uses bracket syntax `[...]` for
    rectangles and `[[...]]` for stadium nodes (no `[N)` ambiguity).
    """
    lines: list = ["graph TB"]

    classifications = report.vba_classifications or []
    if not classifications:
        lines.append('empty["No VBA modules"]')
        return "\n".join(lines)

    # Group module names by classification
    by_class: dict = defaultdict(list)
    for c in classifications:
        by_class[c.inferred_type].append(c.module_name)
    for k in by_class:
        by_class[k].sort(key=str.lower)

    # Stable order: known classes first
    known_order = [
        "data-loader", "transformer", "report-writer",
        "ui-handler", "dead-suspected", "mixed",
    ]
    class_keys = [k for k in known_order if k in by_class] + sorted(
        [k for k in by_class.keys() if k not in known_order]
    )

    # D6: pick the top-2 classes by member count for the visual diagram;
    # the others are summarized in the preceding Markdown table.
    class_keys_sorted_by_size = sorted(
        class_keys, key=lambda c: (-len(by_class[c]), c)
    )
    visible_classes = class_keys_sorted_by_size[:2]

    # Compute how many module nodes we can fit
    # Reserve: 2 class nodes + up to len(visible_classes) "+ N more" placeholders.
    # That leaves _VBA_DIAGRAM_MAX_NODES - 2 * len(visible_classes) for modules
    # (worst case where every class is truncated).
    budget_modules = max(0, _VBA_DIAGRAM_MAX_NODES - 2 * len(visible_classes))
    per_class_quota: dict = {}
    if visible_classes:
        # Distribute budget proportionally to size
        sizes = [len(by_class[c]) for c in visible_classes]
        total = sum(sizes) or 1
        for c, size in zip(visible_classes, sizes):
            per_class_quota[c] = max(1, int(round(budget_modules * size / total)))
        # Adjust so sum == budget
        diff = budget_modules - sum(per_class_quota.values())
        if diff != 0 and visible_classes:
            per_class_quota[visible_classes[0]] += diff

    used: set = set()

    # Class-level nodes
    class_id_map: dict = {}
    for cls in visible_classes:
        cid = _safe_node_id(cls, prefix="C")
        candidate = cid
        i = 1
        while candidate in used:
            candidate = f"{cid}_{i}"
            i += 1
        used.add(candidate)
        class_id_map[cls] = candidate
        # D6 bracket bug fix: use [[ ... ]] for stadium (was buggy `[18)` form
        # because of the `[N]` count inside bracketed labels). New label
        # avoids square brackets in the inner text by using parentheses.
        lbl = f"{cls} ({len(by_class[cls])})"
        lines.append(f'{candidate}[["{_safe_node_label(lbl)}"]]')

    # Module-level nodes (limited per quota)
    mod_id_map: dict = {}
    truncations: dict = {}
    for cls in visible_classes:
        modules_in_class = by_class[cls]
        quota = per_class_quota.get(cls, 0)
        shown = modules_in_class[:quota]
        truncated = len(modules_in_class) - len(shown)
        if truncated > 0:
            truncations[cls] = truncated
        for mod in shown:
            mid = _safe_node_id(mod, prefix="M")
            candidate = mid
            i = 1
            while candidate in used:
                candidate = f"{mid}_{i}"
                i += 1
            used.add(candidate)
            mod_id_map[(cls, mod)] = candidate
            # D6 bracket fix: rectangle node syntax is [...]
            lines.append(f'{candidate}["{_safe_node_label(mod)}"]')

    # Edges class -> module
    for cls in visible_classes:
        modules_in_class = by_class[cls]
        quota = per_class_quota.get(cls, 0)
        shown = modules_in_class[:quota]
        for mod in shown:
            mid = mod_id_map[(cls, mod)]
            lines.append(f"{class_id_map[cls]} --> {mid}")
        if cls in truncations:
            # Add a "+ N more" placeholder edge
            placeholder_id = _safe_node_id(f"more_{cls}", prefix="X")
            cnt = 1
            while placeholder_id in used:
                placeholder_id = _safe_node_id(f"more_{cls}_{cnt}", prefix="X")
                cnt += 1
            used.add(placeholder_id)
            lines.append(f'{placeholder_id}["+ {truncations[cls]} more"]')
            lines.append(f"{class_id_map[cls]} --> {placeholder_id}")

    # Color the visible classifications
    color_map = {
        "data-loader": "#90caf9",
        "transformer": "#a5d6a7",
        "report-writer": "#ffcc80",
        "ui-handler": "#ce93d8",
        "dead-suspected": "#bdbdbd",
        "mixed": "#fff59d",
    }
    for cls, color in color_map.items():
        if cls in class_id_map:
            cls_safe = _safe_node_id(cls, prefix="C")
            lines.append(
                f"classDef cls_{cls_safe} fill:{color},stroke:#333333,color:#000000;"
            )
            lines.append(f"class {class_id_map[cls]} cls_{cls_safe}")

    return "\n".join(lines)


def _build_workflow_mermaid(report) -> str:
    """Diagram (Workflow Sequence) — graph LR of buttons -> Subs -> sheets.

    Capped at <= 15 nodes for readability (D6).
    """
    HARD_CAP = _VBA_DIAGRAM_MAX_NODES  # 15
    lines: list = ["graph LR"]
    wf = (getattr(report, "workflow", None) or {})
    steps = wf.get("steps") or []
    if not steps:
        lines.append('empty["No buttons or event handlers detected"]')
        return "\n".join(lines)
    # Reserve 1 slot for "+ N more" placeholder when we'll truncate.
    will_truncate = len(steps) > 4
    budget = HARD_CAP - (1 if will_truncate else 0)

    used: set = set()
    nodes_added = 0
    steps_emitted = 0
    for s in steps:
        # Each step adds up to 3 new nodes (user, sub, sheet). Stop emitting
        # when adding all 3 would exceed budget.
        if nodes_added + 3 > budget:
            break
        # User node
        user_id = _safe_node_id(f"U_{s.order}", prefix="U")
        if user_id not in used:
            used.add(user_id)
            label = f"User clicks '{(s.label or s.sub_name)[:25]}'"
            lines.append(f'{user_id}["{_safe_node_label(label)}"]')
            nodes_added += 1
        # Sub node
        sub_id = _safe_node_id(f"S_{s.sub_name}_{s.order}", prefix="S")
        if sub_id not in used:
            used.add(sub_id)
            sub_label = f"{s.sub_name} ({s.module_name})"
            lines.append(f'{sub_id}(["{_safe_node_label(sub_label[:40])}"])')
            nodes_added += 1
        lines.append(f"{user_id} --> {sub_id}")
        # First write target as a sheet node
        first_write = s.writes_sheets[0] if s.writes_sheets else (s.sheet or "")
        if first_write:
            sh_id = _safe_node_id(f"SH_{first_write}", prefix="SH")
            if sh_id not in used:
                used.add(sh_id)
                lines.append(f'{sh_id}["{_safe_node_label(first_write[:25])}"]')
                nodes_added += 1
            lines.append(f"{sub_id} -->|writes| {sh_id}")
        steps_emitted += 1

    omitted = len(steps) - steps_emitted
    if omitted > 0:
        lines.append(f'more["+ {omitted} more steps"]:::omitted')
        lines.append("classDef omitted fill:#eeeeee,stroke:#999999,color:#666666;")
    return "\n".join(lines)


def _build_pillar_impact_mermaid(report) -> str:
    """Diagram 3 — graph LR of top-5 pillars to their affected sheets."""
    lines: list = ["graph LR"]

    pillars = report.pillars or []
    if not pillars:
        lines.append('empty["No pillar cells reach the threshold"]')
        return "\n".join(lines)

    top = pillars[:_DIAG3_MAX_PILLARS]
    used: set = set()
    p_ids: dict = {}
    for i, p in enumerate(top, start=1):
        pid = _safe_node_id(p.location, prefix=f"P{i}_")
        c = pid
        n = 1
        while c in used:
            c = f"{pid}_{n}"
            n += 1
        used.add(c)
        p_ids[i - 1] = c
        if p.member_count > 1:
            label = f"{p.location} x{p.member_count}"
        else:
            label = p.location
        lines.append(f'{c}(["{_safe_node_label(label)}"])')

    # Sheet nodes
    sheet_ids: dict = {}
    for p in top:
        for s in p.affected_sheets:
            if s in sheet_ids:
                continue
            sid = _safe_node_id(s, prefix="SH")
            c = sid
            n = 1
            while c in used:
                c = f"{sid}_{n}"
                n += 1
            used.add(c)
            sheet_ids[s] = c
            lines.append(f'{c}["{_safe_node_label(s)}"]')

    # Edges with fan-in label
    for i, p in enumerate(top):
        for s in p.affected_sheets:
            lines.append(f"{p_ids[i]} -->|fan-in {p.fan_in}| {sheet_ids[s]}")

    # Color the pillar nodes
    if p_ids:
        pid_list = ",".join(p_ids[i] for i in sorted(p_ids.keys()))
        lines.append("classDef pillar fill:#ffab91,stroke:#bf360c,color:#000000;")
        lines.append(f"class {pid_list} pillar")
    if sheet_ids:
        sid_list = ",".join(sorted(sheet_ids.values()))
        lines.append("classDef sheet fill:#b3e5fc,stroke:#01579b,color:#000000;")
        lines.append(f"class {sid_list} sheet")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Markdown body — per-section builders for reuse across MD/HTML
# ---------------------------------------------------------------------------

def _section_file_meta(report) -> list:
    L: list = []
    m = report.meta
    L.append("## File metadata")
    L.append("")
    L.append("| Field | Value |")
    L.append("|---|---|")
    L.append(f"| File name | `{m.file_name}` |")
    L.append(f"| File size | {m.file_size_bytes:,} bytes ({m.file_size_bytes/1024:.1f} KB) |")
    L.append(f"| SHA-256 | `{m.sha256}` |")
    if report.sanitized:
        L.append(f"| Sanitize mode | **active** (cell values redacted) |")
    L.append("")
    return L


def _section_basic_stats(report) -> list:
    L: list = []
    b = report.basic_stats
    L.append("## Basic statistics")
    L.append("")
    L.append("| Metric | Value |")
    L.append("|---|---|")
    L.append(f"| Sheet count (total) | {b.sheet_count} |")
    L.append(f"| Sheet count visible / hidden / veryHidden | "
             f"{b.sheet_count_visible} / {b.sheet_count_hidden} / {b.sheet_count_very_hidden} |")
    L.append(f"| Non-empty cells | {b.cell_count_nonempty:,} |")
    L.append(f"| Formula cells | {b.cell_count_formula:,} |")
    L.append(f"| Unique non-formula values | {b.cell_count_unique_values:,} |")
    L.append(f"| Named ranges | {b.named_range_count} |")
    L.append(f"| Conditional formatting rules | {b.conditional_formatting_count} |")
    L.append(f"| Data validation rules | {b.data_validation_count} |")
    L.append(f"| VBA modules | {b.vba_module_count} |")
    L.append(f"| VBA total lines | {b.vba_total_lines:,} |")
    L.append(f"| Cell-level parse errors (logged + skipped) | {b.parse_errors_count} |")
    L.append("")
    return L


def _section_sheets(report) -> list:
    L: list = []
    L.append("## Sheets")
    L.append("")
    L.append("| Sheet | State | Rows | Cols | Non-empty | Formula | Max ref | CF | DV |")
    L.append("|---|---|---|---|---|---|---|---|---|")
    for s in report.sheets:
        L.append(
            f"| `{_md_escape(s.name)}` | {s.state} | {s.rows_used} | {s.cols_used} | "
            f"{s.cells_nonempty} | {s.cells_formula} | {s.max_ref or '—'} | "
            f"{s.conditional_formatting_count} | {s.data_validation_count} |"
        )
    L.append("")
    return L


def _section_named_ranges(report) -> list:
    L: list = []
    L.append("## Named ranges")
    L.append("")
    if not report.named_ranges:
        L.append("_No named ranges defined._")
    else:
        L.append("| Name | Scope | Reference |")
        L.append("|---|---|---|")
        for nr in report.named_ranges:
            L.append(f"| `{_md_escape(nr.name)}` | `{_md_escape(nr.scope)}` | `{_md_escape(nr.ref)}` |")
    L.append("")
    return L


def _section_complexity(report) -> list:
    L: list = []
    c = report.complexity
    L.append("## Complexity score")
    L.append("")
    L.append(f"**Total: {c.total} / 100**")
    L.append("")
    L.append("| Sub-score | Value | Bar | Rationale |")
    L.append("|---|---|---|---|")
    sub = c.sub_scores
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        bar = _sub_score_bar(val)
        L.append(f"| {key.replace('_',' ')} | {val}/20 | `{bar}` | {_md_escape(c.rationale[key])} |")
    L.append("")
    return L


def _format_pillar_value_cell(p) -> str:
    """Inline value display for the new Pillar table column (D1)."""
    v = (getattr(p, "value", "") or "").strip()
    kind = getattr(p, "value_kind", "")
    if not v:
        return "_(empty)_"
    if kind == "formula":
        return f"`{_md_escape(v)}` (formula)"
    if kind == "number":
        return f"`{_md_escape(v)}`"
    if kind == "text":
        return f"`{_md_escape(v)}`"
    return f"`{_md_escape(v)}`"


def _format_pillar_label_cell(p) -> str:
    """Combine row_header/col_header/named_range into one column."""
    parts: list = []
    rh = (getattr(p, "row_header", "") or "").strip()
    ch = (getattr(p, "col_header", "") or "").strip()
    nr = (getattr(p, "named_range", "") or "").strip()
    if nr:
        parts.append(f"named range `{_md_escape(nr)}`")
    if rh:
        parts.append(f"row label `{_md_escape(rh[:25])}`")
    if ch and ch != rh:
        parts.append(f"col header `{_md_escape(ch[:25])}`")
    return "; ".join(parts) if parts else "—"


def _section_pillars(report, top_n: int = None, with_drilldown: bool = True,
                     anchor_id: str = "") -> list:
    """Pillar table — D1 column update: Cell | Value | Label | Members | Fan-in | …

    `top_n`: when set, only render top-N rows (used by Top Impact Findings).
              When None, render all (used by Reference Appendix 8.1).
    """
    L: list = []
    L.append("## Pillar cells — systemic single-points-of-impact")
    L.append("")
    L.append("_Cells with the highest fan-in (most-referenced). Modifying any of these "
             "cascades through many formulas; treat them as critical change points. "
             "Equivalent column cells (same fan-in, same affected sheets) are grouped._")
    L.append("")
    L.append("_**What this means**: each row is a cell whose value flows into "
             "many formulas — change it once, change many calculations at once. "
             "The Value and Label columns answer Michael's #1 ask: \"what IS this cell?\"._")
    L.append("")
    if not report.pillars:
        L.append(f"_No cells reach the pillar threshold "
                 f"(fan-in ≥ {report.methodology['logic_depth_thresholds']['pillar-fanin-min']})._")
        L.append("")
        return L

    rows = report.pillars if top_n is None else report.pillars[:top_n]
    L.append("| Rank | Cell / range | Value | Label | Members | Fan-in | Affected sheets | Kind | Narrative |")
    L.append("|---|---|---|---|---|---|---|---|---|")
    for i, p in enumerate(rows, start=1):
        sheets_str = ", ".join(f"`{s}`" for s in p.affected_sheets[:5])
        if len(p.affected_sheets) > 5:
            sheets_str += f" (+{len(p.affected_sheets) - 5} more)"
        L.append(
            f"| {i} | `{_md_escape(p.location)}` | {_format_pillar_value_cell(p)} | "
            f"{_format_pillar_label_cell(p)} | {p.member_count} | {p.fan_in} | "
            f"{sheets_str} | {p.pillar_kind} | {_md_escape(p.narrative)} |"
        )
    if top_n is not None and len(report.pillars) > top_n:
        L.append("")
        L.append(f"_({len(report.pillars) - top_n} more pillar(s) — see Reference Appendix §8.1.)_")
    L.append("")
    if with_drilldown:
        L.append("**Top-5 pillar drilldown — sample dependents:**")
        L.append("")
        for p in rows[:5]:
            prefix = (f"`{p.location}` ({p.fan_in} dependents per cell × {p.member_count} cells)"
                      if p.member_count > 1
                      else f"`{p.location}` ({p.fan_in} dependents)")
            L.append(f"- **{prefix}**:")
            for d in p.sample_dependents:
                L.append(f"    - `{d}`")
            if p.member_count > 1 and p.member_refs:
                preview = ", ".join(f"`{r}`" for r in p.member_refs[:3])
                extra = f" (+{len(p.member_refs) - 3} more)" if len(p.member_refs) > 3 else ""
                L.append(f"    - _Group members:_ {preview}{extra}")
        L.append("")
    return L


def _section_anomalies(report) -> list:
    L: list = []
    L.append("## Magic-number anomalies — outliers within formula clusters")
    L.append("")
    L.append("_Inside groups of cells that share the same formula shape, this section "
             "flags positions where a small minority uses a different numeric constant. "
             "These are exactly the kind of \"why is this row's discount different?\" "
             "findings that often indicate a missed update or a deliberate (but undocumented) "
             "carve-out._")
    L.append("")
    if not report.anomalies:
        L.append("_No magic-number anomalies detected. Either no large duplicated-formula "
                 "clusters were found, or every cluster's numeric constants are perfectly consistent._")
        L.append("")
        return L
    L.append("| # | Cluster sample | Cluster size | Mode value | Outlier value | Outlier locations | Confidence | Narrative |")
    L.append("|---|---|---|---|---|---|---|---|")
    for i, a in enumerate(report.anomalies, start=1):
        locs = ", ".join(f"`{loc}`" for loc in a.outlier_locations[:5])
        if len(a.outlier_locations) > 5:
            locs += f" (+{len(a.outlier_locations) - 5} more)"
        L.append(
            f"| {i} | `{_md_escape(a.cluster_pattern_sample)}` | "
            f"{a.cluster_size} | `{_md_escape(a.mode_value)}` "
            f"({a.mode_count}/{a.cluster_size}) | "
            f"`{_md_escape(a.outlier_value)}` ({a.outlier_count}/{a.cluster_size}) | "
            f"{locs} | {a.confidence} | {_md_escape(a.narrative)} |"
        )
    L.append("")
    return L


def _section_smells(report) -> list:
    L: list = []
    L.append("## Smells catalog (Hermans 2015)")
    L.append("")
    L.append(f"_{len(report.smells)} smell findings across {len({s.smell_type for s in report.smells})} smell types._")
    L.append("")
    by_type: dict = defaultdict(list)
    for s in report.smells:
        by_type[s.smell_type].append(s)
    for st in SMELL_TYPES:
        items = by_type.get(st, [])
        L.append(f"### `{st}` — {len(items)} findings")
        L.append("")
        if not items:
            L.append("_No findings above threshold._")
            L.append("")
            continue
        L.append("| Location | Metric | Severity | Confidence | Evidence |")
        L.append("|---|---|---|---|---|")
        for s in items[:20]:
            L.append(
                f"| `{_md_escape(s.location)}` | {s.metric:g} | {s.severity} | "
                f"{s.confidence} | {_md_escape(s.evidence)} |"
            )
        if len(items) > 20:
            L.append("")
            L.append(f"_({len(items)-20} more findings of this type — see `audit.json`.)_")
        L.append("")
    return L


def _section_magic_index(report) -> list:
    L: list = []
    L.append("## Magic-number index (top 20)")
    L.append("")
    if not report.magic_numbers:
        L.append("_No non-trivial numeric literals found._")
    else:
        L.append("| Value | Count | First location | Source | Sample context |")
        L.append("|---|---|---|---|---|")
        for mn in report.magic_numbers:
            L.append(
                f"| `{_md_escape(mn.value)}` | {mn.occurrence_count} | "
                f"`{_md_escape(mn.first_location)}` | {mn.location_kind} | "
                f"`{_md_escape(mn.sample_context)}` |"
            )
    L.append("")
    return L


def _section_vba(report) -> list:
    L: list = []
    L.append("## VBA modules + classification")
    L.append("")
    if not report.vba_modules:
        L.append("_No VBA modules found._")
        L.append("")
        return L
    cls_by_name = {c.module_name: c for c in (report.vba_classifications or [])}
    L.append("| Module | Type | LOC | #Sub | #Func | Inferred type | Confidence | Reads | Writes | Ext calls | OnErrorResumeNext |")
    L.append("|---|---|---|---|---|---|---|---|---|---|---|")
    for vm in report.vba_modules:
        n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
        n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
        cls = cls_by_name.get(vm.name)
        inferred = cls.inferred_type if cls else "—"
        cls_conf = cls.confidence if cls else "—"
        reads = ", ".join(f"`{s}`" for s in (cls.reads_sheets if cls else [])) or "—"
        writes = ", ".join(f"`{s}`" for s in (cls.writes_sheets if cls else [])) or "—"
        ext_call = ("yes" if (cls and cls.external_calls) else "no") if cls else "—"
        L.append(
            f"| `{_md_escape(vm.name)}` | {vm.type} | {vm.line_count} | "
            f"{n_sub} | {n_func} | **{inferred}** | {cls_conf} | "
            f"{_md_escape(reads)} | {_md_escape(writes)} | {ext_call} | "
            f"{'yes' if vm.has_on_error_resume_next else 'no'} |"
        )
    L.append("")
    L.append("**VBA module details — external keywords & range literals:**")
    L.append("")
    L.append("| Module | External keywords | Range literals (uniq) | Classifier rationale |")
    L.append("|---|---|---|---|")
    for vm in report.vba_modules:
        ext = ", ".join(vm.external_keywords) or "—"
        ranges = len(vm.range_literals)
        cls = cls_by_name.get(vm.name)
        rationale = cls.rationale if cls else "—"
        L.append(
            f"| `{_md_escape(vm.name)}` | {_md_escape(ext)} | {ranges} | "
            f"{_md_escape(rationale)} |"
        )
    L.append("")
    return L


def _section_risks(report) -> list:
    L: list = []
    r = report.risk_indicators
    L.append("## Risk indicators")
    L.append("")
    L.append("| Indicator | Value |")
    L.append("|---|---|")
    L.append(f"| Hidden sheets | {len(r.hidden_sheets)} (`{', '.join(r.hidden_sheets) if r.hidden_sheets else '—'}`) |")
    L.append(f"| Very-hidden sheets | {len(r.very_hidden_sheets)} (`{', '.join(r.very_hidden_sheets) if r.very_hidden_sheets else '—'}`) |")
    L.append(f"| Cross-sheet referencing formulas | {r.cross_sheet_reference_count} |")
    L.append(f"| Cells with formula errors (cached) | {len(r.cells_with_errors)} |")
    L.append(f"| External workbook reference patterns | {len(r.external_workbook_references)} |")
    L.append(f"| Circular reference suspects | {len(r.circular_reference_suspects)} |")
    L.append(f"| Parse errors logged | {len(r.parse_errors)} |")
    L.append("")
    if r.cells_with_errors:
        L.append("**Formula error cells (first 20):**")
        L.append("")
        L.append("| Sheet | Ref | Error |")
        L.append("|---|---|---|")
        for ec in r.cells_with_errors[:20]:
            L.append(f"| `{_md_escape(ec['sheet'])}` | `{ec['ref']}` | `{ec['error_token']}` |")
        L.append("")
    if r.external_workbook_references:
        L.append("**External workbook references:**")
        L.append("")
        for ext in r.external_workbook_references[:20]:
            L.append(f"- `{_md_escape(ext)}`")
        L.append("")
    if r.circular_reference_suspects:
        L.append("**Circular reference suspects (first 20):**")
        L.append("")
        for cs in r.circular_reference_suspects[:20]:
            L.append(f"- `{_md_escape(cs)}`")
        L.append("")
    return L


def _section_diagrams(report) -> list:
    L: list = []
    L.append("## Sheet data flow diagram")
    L.append("")
    L.append("_Cross-sheet formula references. Edge label = number of formulas with cross-sheet refs from source -> target sheet. "
             "Yellow = hidden, red = veryHidden, default = visible._")
    L.append("")
    L.append("```mermaid")
    L.append(_build_sheet_dataflow_mermaid(report))
    L.append("```")
    L.append("")

    L.append("## VBA classification overview diagram")
    L.append("")
    L.append("_Modules grouped under their inferred type._")
    L.append("")
    L.append("```mermaid")
    L.append(_build_vba_classification_mermaid(report))
    L.append("```")
    L.append("")

    L.append("## Pillar impact diagram")
    L.append("")
    L.append(f"_Top-{_DIAG3_MAX_PILLARS} pillar cells and the sheets they cascade into._")
    L.append("")
    L.append("```mermaid")
    L.append(_build_pillar_impact_mermaid(report))
    L.append("```")
    L.append("")
    return L


def _section_methodology(report) -> list:
    L: list = []
    L.append("## Methodology")
    L.append("")
    libs = report.methodology["library_versions"]
    L.append(f"- Engine: `openpyxl` {libs['openpyxl']} (cells/structure), "
             f"`oletools.olevba` {libs['oletools']} (VBA), "
             f"`formulas` {libs['formulas']} (tokenizer only — no evaluator).")
    thr = report.methodology["smell_thresholds"]
    L.append(
        f"- Smell thresholds: multiple-references ≥ {thr['multiple-references']}; "
        f"long-calculation-chain depth ≥ {thr['long-calculation-chain']}; "
        f"conditional-complexity nesting ≥ {thr['conditional-complexity']}; "
        f"multiple-operations ≥ {thr['multiple-operations']}; "
        f"duplicated-formulas pattern frequency ≥ {thr['duplicated-formulas']}."
    )
    ld = report.methodology["logic_depth_thresholds"]
    L.append(
        f"- Logic-depth thresholds: pillar fan-in ≥ {ld['pillar-fanin-min']} (top {ld['pillar-top-n']} after dedupe); "
        f"anomaly cluster size ≥ {ld['anomaly-cluster-min-size']}, outlier fraction ≤ {ld['anomaly-outlier-fraction']}."
    )
    L.append("- Pillar dedupe: cells in the same column with identical fan-in and identical "
             "affected-sheet set are collapsed into one column-block entry. After dedupe we surface the top "
             f"{ld['pillar-top-n']} distinct entries.")
    L.append("- VBA classifier categories: " + ", ".join(f"`{c}`" for c in report.methodology["vba_classifier_categories"]) + ".")
    L.append("- Confidence semantics: `high` = exact / deterministic count; "
             "`medium` = tokenizer-based with well-defined rules; "
             "`low` = statistical inference or analysis skipped due to scale.")
    L.append("- Domain hint detector: pure keyword matching (case-insensitive, word-boundary) "
             "against sheet names, named ranges, and VBA Sub/Function names. No LLM, no inference.")
    L.append("- Trivial numbers excluded from magic-number index: " +
             ", ".join(f"`{n}`" for n in sorted(TRIVIAL_NUMBERS)) + ".")
    L.append("- Reliability contract: same input → byte-identical `audit.md`, `audit.json`, and `audit.html`. "
             "All outputs use sorted dict keys and lexicographic list ordering. No timestamps. No CDN fetch at audit time.")
    if report.sanitized:
        L.append("- **Sanitize mode active**: every non-formula cell value has been replaced with "
                 "`<redacted>` before any analysis ran. Formulas, VBA source, smells, pillars, anomalies, "
                 "and structural counts remain accurate. The SHA-256 above is of the original file.")
    L.append("")
    L.append("---")
    L.append("")
    L.append("_End of report._")
    return L


# ---------------------------------------------------------------------------
# Round-3 new section builders — Workflow / Data Flow / VBA Walkthrough /
# Domain Findings / Top Impact Findings / Glossary.
# ---------------------------------------------------------------------------

def _section_workflow_guide(report) -> list:
    """D3: Workflow Guide — operational walkthrough.

    Derived from xl/drawings + VBA structure. When NO buttons + NO event
    handlers, the section reports "no buttons detected" gracefully.
    """
    L: list = []
    L.append("## Workflow Guide")
    L.append("")
    L.append("_Inferred from the workbook's VBA structure and embedded form-control "
             "buttons (no AI). This is how the workbook is operationally used._")
    L.append("")
    L.append("_**What this means**: the steps below describe the sequence a user "
             "follows — typically: open the workbook, enter inputs, click button(s), "
             "read the resulting cells. This is structural inference; semantic "
             "narrative (e.g. \"this calculates capacity utilization\") would require "
             "Track B LLM augmentation._")
    L.append("")
    wf = (getattr(report, "workflow", None) or {})
    steps = wf.get("steps") or []

    if not steps:
        L.append("_No user-callable buttons or event handlers detected — this "
                 "workbook appears to be formula-driven only (no macro entry points "
                 "found in `xl/drawings/*` or VBA event subs). Users would interact "
                 "with the workbook by editing input cells and reading formula "
                 "outputs directly._")
        L.append("")
        return L

    for s in steps:
        L.append(f"<!-- LLM-AUGMENT: workflow-step:{s.order} -->")
        L.append(f"**Step {s.order}**: User clicks button **{_md_escape(s.label or s.sub_name)}** "
                 f"on sheet `{_md_escape(s.sheet)}` (bound to "
                 f"`{_md_escape(s.module_name)}.{_md_escape(s.sub_name)}`).")
        if s.reads_sheets:
            L.append(f"- Reads: " + ", ".join(f"`{_md_escape(r)}`" for r in s.reads_sheets[:5])
                     + (f" (+{len(s.reads_sheets)-5} more)" if len(s.reads_sheets) > 5 else ""))
        if s.writes_sheets:
            L.append(f"- Writes: " + ", ".join(f"`{_md_escape(w)}`" for w in s.writes_sheets[:5])
                     + (f" (+{len(s.writes_sheets)-5} more)" if len(s.writes_sheets) > 5 else ""))
        if s.calls:
            L.append(f"- Calls: " + ", ".join(f"`{_md_escape(c)}`" for c in s.calls[:5])
                     + (f" (+{len(s.calls)-5} more)" if len(s.calls) > 5 else ""))
        L.append("")

    # Mermaid sequence diagram
    L.append("**Workflow sequence diagram**:")
    L.append("")
    L.append("```mermaid")
    L.append(_build_workflow_mermaid(report))
    L.append("```")
    L.append("")
    return L


def _section_data_flow_story(report) -> list:
    """D5: Data Flow Story — per-sheet prose paragraph (top 8 by density).

    Renders BEFORE the existing Mermaid sheet-flow diagram (which becomes
    "evidence" for this prose).
    """
    L: list = []
    L.append("## Data Flow Story")
    L.append("")
    L.append("_Plain-language description of how data flows between sheets, before "
             "the schematic diagram. Derived from formula cross-sheet references and "
             "VBA write targets._")
    L.append("")
    L.append("_**What this means**: each paragraph below tells you whether a sheet "
             "is **input** (user types here), **derived** (populated by formulas or "
             "macros), or **mixed** — and where its values come from / go to._")
    L.append("")

    # Sort sheets by combined density: cells_nonempty + cells_formula
    sheets_ranked = sorted(
        report.sheets,
        key=lambda s: -(s.cells_nonempty + s.cells_formula * 2),
    )
    top_sheets = sheets_ranked[:8]

    # Build incoming/outgoing maps from _sheet_edges
    edges = getattr(report, "_sheet_edges", None) or {}
    incoming: dict = defaultdict(list)
    outgoing: dict = defaultdict(list)
    for (src, tgt), cnt in edges.items():
        incoming[tgt].append((src, cnt))
        outgoing[src].append((tgt, cnt))
    for d in (incoming, outgoing):
        for k in d:
            d[k].sort(key=lambda x: (-x[1], x[0]))

    # Build VBA-write set (from classifications)
    vba_writes_by_sheet: dict = defaultdict(list)
    for c in (report.vba_classifications or []):
        for w_sheet in c.writes_sheets:
            vba_writes_by_sheet[w_sheet].append(c.module_name)

    # Pillar count per sheet
    pillar_count_by_sheet: dict = defaultdict(int)
    for p in (report.pillars or []):
        # Extract source sheet from location string "Sheet!REF..."
        if "!" in p.location:
            src_sheet = p.location.split("!", 1)[0]
            pillar_count_by_sheet[src_sheet] += p.member_count

    # Per-sheet error cells
    err_cells_by_sheet: dict = defaultdict(list)
    for ec in report.risk_indicators.cells_with_errors:
        err_cells_by_sheet[ec["sheet"]].append(ec)

    for s in top_sheets:
        L.append(f"<!-- LLM-AUGMENT: data-flow:{s.name} -->")
        L.append(f"### `{_md_escape(s.name)}` ({s.state}, "
                 f"{s.rows_used} rows × {s.cols_used} cols, "
                 f"{s.cells_nonempty} non-empty cells)")
        # Role inference
        n_in = sum(c for _, c in incoming.get(s.name, []))
        n_out = sum(c for _, c in outgoing.get(s.name, []))
        n_form = s.cells_formula
        n_vba_writes = len(vba_writes_by_sheet.get(s.name, []))

        if n_form == 0 and n_in == 0 and n_vba_writes == 0:
            role = ("**input sheet** — no formulas, no inbound cross-sheet "
                    "references, no VBA writes. Likely user-driven manual entry.")
        elif n_form == 0 and (n_in > 0 or n_vba_writes > 0):
            role = ("**derived sheet (no formulas)** — populated by VBA macros "
                    "or referenced as a source by other sheets but contains only values.")
        elif n_form > 0 and n_out > n_in:
            role = ("**aggregator/output sheet** — many of its formulas pull from "
                    "other sheets; downstream usage limited.")
        elif n_form > 0:
            role = ("**computed sheet** — populated by formulas with cross-sheet "
                    "lookups.")
        else:
            role = "**mixed**."
        L.append(f"**Role**: {role}")

        # Sources / consumers
        if outgoing.get(s.name):
            srcs = ", ".join(f"`{src}` ({cnt})" for src, cnt in outgoing[s.name][:4])
            L.append(f"**Sources** (formulas in this sheet read from): {srcs}.")
        if incoming.get(s.name):
            tgts = ", ".join(f"`{tgt}` ({cnt})" for tgt, cnt in incoming[s.name][:4])
            L.append(f"**Consumers** (other sheets reading this one): {tgts}.")
        if n_vba_writes > 0:
            mods = vba_writes_by_sheet[s.name][:3]
            mods_str = ", ".join(f"`{m}`" for m in mods)
            extra = f" (+{n_vba_writes-3} more)" if n_vba_writes > 3 else ""
            L.append(f"**VBA writes to this sheet**: {mods_str}{extra}.")
        if pillar_count_by_sheet.get(s.name):
            L.append(f"**Pillar cells**: {pillar_count_by_sheet[s.name]} cell(s) "
                     "in this sheet are pillars (high fan-in change points).")
        if err_cells_by_sheet.get(s.name):
            ec = err_cells_by_sheet[s.name][:1][0]
            L.append(f"**Manual override risk**: cell `{ec['ref']}` has cached error "
                     f"`{ec['error_token']}` — see Risk indicators §8.6.")
        L.append("")

    if len(report.sheets) > len(top_sheets):
        L.append(f"_({len(report.sheets) - len(top_sheets)} more sheet(s) — full table in §8.3.)_")
        L.append("")

    # The schematic diagram
    L.append("**Sheet data flow diagram**:")
    L.append("")
    L.append("_Cross-sheet formula references. Edge label = number of formulas "
             "with cross-sheet refs from source -> target sheet. "
             "Yellow = hidden, red = veryHidden._")
    L.append("")
    L.append("```mermaid")
    L.append(_build_sheet_dataflow_mermaid(report))
    L.append("```")
    L.append("")
    return L


def _section_top_impact_findings(report) -> list:
    """Top Impact Findings — Top-5 each of pillars / anomalies / smells / risks.

    SHORT and READABLE — appendix has the full data.
    """
    L: list = []
    L.append("## Top Impact Findings")
    L.append("")
    L.append("_Top-N filtered list across pillars, anomalies, smells, and risks. "
             "Full catalogs live in the Reference Appendix (§8)._")
    L.append("")

    # Pillars (top 5)
    L.append("### Top-5 Pillar Cells (single-points-of-impact)")
    L.append("")
    if not report.pillars:
        L.append("_No pillar cells detected._")
    else:
        L.append("| # | Cell / range | Value | Label | Fan-in | Affected sheets |")
        L.append("|---|---|---|---|---|---|")
        for i, p in enumerate(report.pillars[:5], start=1):
            sheets_str = ", ".join(f"`{s}`" for s in p.affected_sheets[:3])
            if len(p.affected_sheets) > 3:
                sheets_str += f" (+{len(p.affected_sheets)-3})"
            L.append(
                f"| {i} | `{_md_escape(p.location)}` | {_format_pillar_value_cell(p)} | "
                f"{_format_pillar_label_cell(p)} | {p.fan_in} | {sheets_str} |"
            )
        L.append("")
        L.append("_See §8.1 for the full pillar table; each row's narrative explains "
                 "the cell's role and impact._")
    L.append("")

    # Anomalies (top 5)
    L.append("### Top-5 Magic-number Anomalies (cluster outliers)")
    L.append("")
    if not report.anomalies:
        L.append("_No magic-number anomalies detected. Either no large duplicated-formula "
                 "clusters, or every cluster's numbers are perfectly consistent._")
    else:
        L.append("| # | Cluster size | Mode | Outlier | Outlier locations | Confidence |")
        L.append("|---|---|---|---|---|---|")
        for i, a in enumerate(report.anomalies[:5], start=1):
            locs = ", ".join(f"`{loc}`" for loc in a.outlier_locations[:3])
            if len(a.outlier_locations) > 3:
                locs += f" (+{len(a.outlier_locations)-3} more)"
            L.append(
                f"| {i} | {a.cluster_size} | `{_md_escape(a.mode_value)}` "
                f"({a.mode_count}/{a.cluster_size}) | "
                f"`{_md_escape(a.outlier_value)}` ({a.outlier_count}/{a.cluster_size}) | "
                f"{locs} | {a.confidence} |"
            )
    L.append("")

    # Top smells (top 5 by metric)
    L.append("### Top-5 Smell Findings")
    L.append("")
    L.append("_**What this means**: code-smell categories from Hermans 2015. "
             "Not bugs — patterns that often indicate maintainability risk._")
    L.append("")
    if not report.smells:
        L.append("_No smells above threshold._")
    else:
        # Dedupe by (smell_type, metric, sheet) so the Top-5 surfaces distinct
        # findings rather than 5 cells from the same column with identical metric.
        # The full unfiltered catalog still lives in §8.2.
        seen_keys: set = set()
        top_smells = []
        for s in sorted(report.smells, key=lambda s: -s.metric):
            sheet = s.location.split('!', 1)[0] if '!' in s.location else s.location
            dedupe_key = (s.smell_type, s.metric, sheet)
            if dedupe_key in seen_keys:
                continue
            seen_keys.add(dedupe_key)
            top_smells.append(s)
            if len(top_smells) >= 5:
                break
        L.append("| # | Type | Location | Metric | Severity | Evidence |")
        L.append("|---|---|---|---|---|---|")
        for i, s in enumerate(top_smells, start=1):
            L.append(
                f"| {i} | `{s.smell_type}` | `{_md_escape(s.location)}` | "
                f"{s.metric:g} | {s.severity} | {_md_escape(s.evidence)} |"
            )
        L.append("")
        L.append("_Full smell catalog: §8.2._")
    L.append("")

    # Top risks
    L.append("### Top Risk Indicators")
    L.append("")
    r = report.risk_indicators
    risk_lines: list = []
    if r.very_hidden_sheets:
        risk_lines.append(f"- **{len(r.very_hidden_sheets)} veryHidden sheet(s)**: "
                          f"`{', '.join(r.very_hidden_sheets)}` — invisible in Excel "
                          f"UI even via Hide/Unhide menu; only VBA reveals them.")
    if r.hidden_sheets:
        risk_lines.append(f"- **{len(r.hidden_sheets)} hidden sheet(s)**: "
                          f"`{', '.join(r.hidden_sheets)}` — not visible by default "
                          f"but can be unhidden via right-click.")
    if r.cells_with_errors:
        risk_lines.append(f"- **{len(r.cells_with_errors)} cell(s) with cached "
                          f"formula errors** (e.g. #REF!, #N/A) — see §8.6.")
    if r.external_workbook_references:
        risk_lines.append(f"- **{len(r.external_workbook_references)} external workbook "
                          f"reference pattern(s)** — fragile cross-file links.")
    if r.circular_reference_suspects:
        risk_lines.append(f"- **{len(r.circular_reference_suspects)} circular "
                          f"reference suspect(s)** — see §8.6.")
    if not risk_lines:
        L.append("_No high-priority risk indicators._")
    else:
        for line in risk_lines:
            L.append(line)
    L.append("")
    return L


def _section_vba_walkthrough(report) -> list:
    """D4: VBA Module Walkthrough — prose narration in call order.

    Each module gets a paragraph; modules with no callers flagged as
    "possibly dead code". LLM-AUGMENT markers reserve the slot for Track B.
    """
    L: list = []
    L.append("## VBA Module Walkthrough")
    L.append("")
    L.append("_Per-module heuristic narrative, ordered by call-graph dependency from "
             "user-callable entry points (button-bound + event handlers). "
             "**This is structural narration, not semantic** — we report what reads/"
             "writes and what calls what, but not what the code MEANS for the business "
             "(that's Track B / LLM augmentation)._")
    L.append("")
    L.append("_**What this means**: each module gets a 4-line summary: structural role, "
             "what it does (sheets it reads/writes, modules it calls), notable patterns "
             "(error handling, magic numbers, loops), and call relationships._")
    L.append("")
    narratives = report.vba_narratives or []
    if not narratives:
        L.append("_No VBA modules to narrate._")
        L.append("")
        return L

    # Group: reachable-from-entry first, then unreachable
    reachable = [n for n in narratives if n.reachable_from_entry]
    unreachable = [n for n in narratives if not n.reachable_from_entry]

    # Cap reachable to top 8 to keep section readable; rest summarized
    REACHABLE_MAX = 8

    # When NO buttons/events were found, we still want SOME narrative shown —
    # otherwise the section is empty for formula-only workbooks. Pick the top
    # modules by line count (likely the most consequential code).
    if not reachable and unreachable:
        # Promote top-N unreachable modules (sorted by LOC desc) to "shown"
        ordered = sorted(unreachable, key=lambda n: -n.line_count)
        narratives_to_show = ordered[:REACHABLE_MAX]
        truly_dead = [n for n in unreachable if n not in narratives_to_show]
        L.append("_No buttons or event handlers detected, so no call-graph entry "
                 "point — modules below are ranked by LOC (a structural proxy "
                 "for likely importance). Each module is flagged as 'possibly "
                 "dead code' since static analysis can't confirm reachability "
                 "without an entry point._")
        L.append("")
        L.append("_**Caveat for many-dead-modules workbooks**: when a large fraction "
                 "of modules are flagged 'possibly dead', the VBA may be inherited "
                 "from another project (donor code), left from prior workbook "
                 "iterations, or genuinely orphaned. Static analysis cannot "
                 "disambiguate — Track B (LLM-augmented) reads the code to tell. "
                 "For now, treat 'possibly dead' as 'no entry point detected', "
                 "not 'definitely unused'._")
        L.append("")
        for narr in narratives_to_show:
            L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
            L.append(f"### `{_md_escape(narr.module_name)}` "
                     f"({narr.inferred_type}, {narr.line_count} lines)")
            L.append("")
            L.append(narr.narrative)
            L.append("")
        if truly_dead:
            L.append(f"### Possibly dead code ({len(truly_dead)} additional module(s))")
            L.append("")
            sample = truly_dead[:8]
            for narr in sample:
                L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
                L.append(f"- `{_md_escape(narr.module_name)}` "
                         f"({narr.inferred_type}, {narr.line_count} lines, "
                         f"{narr.sub_count + narr.func_count} subs/funcs)")
            if len(truly_dead) > len(sample):
                L.append(f"- _(+{len(truly_dead) - len(sample)} more — see §8.7.)_")
            L.append("")
        return L

    reachable_show = reachable[:REACHABLE_MAX]
    for narr in reachable_show:
        L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
        L.append(f"### `{_md_escape(narr.module_name)}` "
                 f"({narr.inferred_type}, {narr.line_count} lines)")
        L.append("")
        L.append(narr.narrative)
        L.append("")

    if len(reachable) > REACHABLE_MAX:
        L.append(f"_({len(reachable) - REACHABLE_MAX} more reachable module(s) — "
                 f"see Reference Appendix §8.7 VBA modules table for full list.)_")
        L.append("")

    if unreachable:
        L.append(f"### Possibly dead code ({len(unreachable)} module(s))")
        L.append("")
        L.append("_The following modules are not reached from any detected button "
                 "or event handler. They may be legacy code, helper libraries pulled "
                 "in but unused, or detection misses (ActiveX controls, dynamic VBA "
                 "calls). Audit before deleting._")
        L.append("")
        sample = unreachable[:8]
        for narr in sample:
            L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
            L.append(f"- `{_md_escape(narr.module_name)}` "
                     f"({narr.inferred_type}, {narr.line_count} lines, "
                     f"{narr.sub_count + narr.func_count} subs/funcs)")
        if len(unreachable) > len(sample):
            L.append(f"- _(+{len(unreachable) - len(sample)} more — see §8.7.)_")
        L.append("")
    return L


def _section_domain_findings(report) -> list:
    """D7: Domain-Specific Findings — only present when ≥1 template fired."""
    L: list = []
    matches = (report.domain_template_matches or [])
    # Filter to high/medium confidence; low-confidence is noise here
    matches = [m for m in matches if m.confidence in ("high", "medium")]
    if not matches:
        return []  # caller skips section

    L.append("## Domain-Specific Findings")
    L.append("")
    L.append("_Detected domain templates (e.g. manufacturing/capacity-planning, "
             "logistics/routing). For each, we cross-checked the workbook against "
             "industry-known hardcoded-risk constants and scheduling-method hallmarks._")
    L.append("")
    L.append("_**What this means**: domain templates pre-populate \"things to look "
             "for\" specific to a vertical. They're a heuristic checklist, not a "
             "verdict — every hit deserves a closer look but isn't necessarily a bug._")
    L.append("")

    for m in matches:
        L.append(f"<!-- LLM-AUGMENT: domain-method:{m.template_key} -->")
        L.append(f"### {m.business_friendly_name}  "
                 f"_(confidence: {m.confidence})_")
        L.append("")
        if m.matched_keywords:
            L.append("**Matched keywords**: " + ", ".join(f"`{k}`" for k in m.matched_keywords) + ".")
            L.append("")

        if m.sheet_role_hits:
            L.append("**Expected sheet roles found in this workbook**:")
            for role, sheets in m.sheet_role_hits:
                L.append(f"- `{role}` — present as {sheets}")
            L.append("")
        if m.sheet_role_misses:
            L.append("**Expected sheet roles NOT found**:")
            for role in m.sheet_role_misses[:5]:
                L.append(f"- `{role}` — no sheet name matches; confirm whether absent or named differently")
            L.append("")

        if m.hardcode_risk_hits:
            L.append("**Common hardcode risks (potential hits)**:")
            for label, evidence in m.hardcode_risk_hits:
                L.append(f"- **{label}** — evidence: {evidence}")
            L.append("")
        else:
            L.append("**Common hardcode risks**: no clear matches found in pillars / "
                     "magic-number index. Either the risks aren't present or are "
                     "named differently.")
            L.append("")

        if m.method_hits:
            L.append("**Scheduling methods detected (hallmark keywords)**:")
            for label, evidence in m.method_hits:
                L.append(f"- **{label}** — evidence: {evidence}")
            L.append("")
    return L


_GLOSSARY_TERMS = {
    "anomaly (magic-number)": (
        "Inside a cluster of cells sharing the same formula shape, a position "
        "where a small minority uses a different numeric constant. Often "
        "signals a missed update or undocumented carve-out."
    ),
    "audit": (
        "The full report this tool produces — markdown + JSON + HTML — for one "
        "xlsm workbook. Pure static analysis, zero LLM, zero network."
    ),
    "BYOA (Bring Your Own AI)": (
        "Distribution model where the customer uses their own LLM "
        "subscription. We do not call any LLM; we just package context "
        "for them to paste into Copilot / Claude / etc."
    ),
    "column-block (pillar)": (
        "When N cells in the same column share the same fan-in count and the "
        "same set of dependent sheets, we collapse them into one pillar entry "
        "with `member_count = N` instead of N separate rows."
    ),
    "complexity score": (
        "A 0-100 composite of 5 sub-scores (data scale, formula depth, "
        "metadata complexity, smell density, VBA mass). Higher = harder to "
        "refactor."
    ),
    "confidence (high/medium/low)": (
        "How certain we are about a finding. `high` = exact deterministic "
        "count; `medium` = tokenizer-based with well-defined rules; "
        "`low` = statistical inference or analysis skipped."
    ),
    "data-loader / transformer / report-writer / ui-handler": (
        "VBA module classifications based on Sub/Function naming patterns "
        "(`Load*`, `Calc*`, `Print*`, `*_Click`) and structural counts "
        "(reads/writes/loops). Heuristic only."
    ),
    "dead-suspected (VBA)": (
        "A module with no Sub calls, no value writes, no name signals — "
        "likely empty or unused. May still be reachable through dynamic "
        "VBA calls; verify before deleting."
    ),
    "domain": (
        "An industry vertical we detected via keyword matching: "
        "capacity-planning / inventory-supply-chain / logistics-routing / "
        "operations-s&op / actuarial-insurance / financial-modeling. "
        "Triggers domain-specific sections when matched."
    ),
    "duplicated-formulas": (
        "Cells whose normalized formula pattern appears in many places. Often "
        "legitimate (column-fill), but can hide outliers — see Magic-number "
        "Anomalies."
    ),
    "fan-in": (
        "How many distinct formulas reference a given cell. Higher fan-in = "
        "modifying that cell ripples through more calculations."
    ),
    "formula relay (pillar)": (
        "A pillar cell that itself is a formula — others read its result, "
        "and it is itself derived. Modifying the formula changes downstream "
        "derivations."
    ),
    "Hermans smells": (
        "Spreadsheet code-smell catalog from Felienne Hermans's 2015 paper "
        "(multiple-references, conditional-complexity, multiple-operations, "
        "duplicated-formulas, magic-numbers, long-calculation-chain)."
    ),
    "incoming / outgoing (cell)": (
        "Cell A's `incoming` set is every cell whose formula reads A. "
        "Cell A's `outgoing` set is every cell A's formula reads. Together "
        "they form the cell-level dataflow graph."
    ),
    "LLM-AUGMENT marker": (
        "An HTML comment like `<!-- LLM-AUGMENT: vba-narration:Module1 -->` "
        "marking a section a future Track B (LLM) ingest step can replace "
        "with richer prose. Invisible in rendered Markdown."
    ),
    "magic number": (
        "A non-trivial numeric literal (not 0/1/2/10/100/-1) embedded "
        "directly in formula or VBA code. Usually a candidate for "
        "extraction into a named constant."
    ),
    "On Error Resume Next": (
        "VBA directive that silently skips the next failing line. "
        "Risky pattern: errors are suppressed, code keeps running with "
        "potentially-bad state. Always reviewed in the audit."
    ),
    "pillar cell": (
        "A cell with high fan-in (≥ 20 by default) — modifying it cascades "
        "to many formulas. The audit's primary diagnostic for \"what "
        "should I never accidentally change?\""
    ),
    "Sanitize mode": (
        "Optional `--sanitize` flag that replaces every non-formula cell "
        "value with `<redacted>` before any analysis. Formulas, VBA "
        "structure, and counts are preserved. Designed for sharing the "
        "audit without leaking the workbook's data."
    ),
    "smell": (
        "A code pattern the audit flags as worth a second look. Not a bug — "
        "a heuristic 'smelly' indicator borrowed from software engineering."
    ),
    "tier (1/1.5/2/3/4/5)": (
        "Our product tiers — Tier 1 = this audit (free, zero-LLM); "
        "Tier 1.5 = LLM-assisted comprehension (BYOA); Tier 2 = exec risk "
        "report; Tier 3 = refactored Python prototype; etc."
    ),
    "Track A / Track B": (
        "Two-track architecture for Tier 1 reports. Track A is fully "
        "static (no LLM); Track B is BYOA — tool produces a dossier + "
        "mega-prompt the user pastes into their own Copilot / Claude, "
        "then pastes the response back."
    ),
    "veryHidden sheet": (
        "Excel hides three states: visible / hidden / veryHidden. "
        "veryHidden cannot be unhidden through the right-click menu — only "
        "VBA can reveal it. Often used for internal config or audit-trail "
        "data the user shouldn't touch."
    ),
    "workflow step": (
        "One operational step a user performs — typically clicking a "
        "button or triggering an event handler. Derived from xl/drawings "
        "+ VBA static analysis, then topologically sorted by sheet write/"
        "read dependency."
    ),
}


def _section_glossary(report) -> list:
    """D8: Glossary — alphabetical plain-language definitions."""
    L: list = []
    L.append("## Glossary")
    L.append("")
    L.append("_Plain-language definitions for terms used throughout this report. "
             "Alphabetical._")
    L.append("")
    for term in sorted(_GLOSSARY_TERMS.keys(), key=str.lower):
        defn = _GLOSSARY_TERMS[term]
        L.append(f"- **{_md_escape(term)}** — {_md_escape(defn)}")
    L.append("")
    return L


def _section_executive_summary_round3(report) -> list:
    """Round-3 executive summary — 3-5 manager-readable headlines.

    Replaces the prior more verbose exec summary with a tightly-filtered
    pyramid top.
    """
    L: list = []
    c = report.complexity
    domain = getattr(report, "domain_hint", None)

    L.append("## Executive Summary")
    L.append("")
    # Headline: complexity + plain-language rendition
    if c.total >= 80:
        complexity_tier = "top tier — substantial refactor effort"
    elif c.total >= 50:
        complexity_tier = "moderately complex — meaningful refactor effort"
    elif c.total >= 20:
        complexity_tier = "manageable — straightforward to read"
    else:
        complexity_tier = "small / lightly-used"
    L.append(f"- **Complexity score: {c.total} / 100** — {complexity_tier}.")

    # Top pillar
    if report.pillars:
        p = report.pillars[0]
        if p.member_count > 1:
            L.append(f"- **Most-referenced cell group**: `{p.location}` "
                     f"({p.member_count} cells, fan-in {p.fan_in}). "
                     f"Each cell drives {p.fan_in} calculations.")
        else:
            value_phrase = ""
            if p.value:
                value_phrase = f" (value `{p.value[:30]}`)"
            L.append(f"- **Single most-impactful cell**: `{p.location}`{value_phrase} "
                     f"feeds {p.fan_in} formulas across "
                     f"{p.affected_sheet_count} sheet(s). Change it, change them all.")

    # Top smell or anomaly
    if report.anomalies:
        a = report.anomalies[0]
        L.append(f"- **Top data-anomaly**: in a cluster of {a.cluster_size} "
                 f"similar formulas, value `{a.outlier_value}` deviates from "
                 f"the norm `{a.mode_value}` at `{a.outlier_locations[0]}` "
                 f"(confidence: {a.confidence}).")
    elif report.smells:
        top_smell = max(report.smells, key=lambda s: s.metric)
        L.append(f"- **Top smell**: `{top_smell.smell_type}` at "
                 f"`{top_smell.location}` (metric={top_smell.metric:g}, "
                 f"severity={top_smell.severity}).")

    # Domain
    if domain is not None and domain.domain != "unknown":
        kw_str = ", ".join(domain.matched_keywords[:5])
        L.append(f"- **Detected domain**: `{domain.domain}` "
                 f"_(confidence: {domain.confidence}, matched: {kw_str}_).")

    # Workflow
    wf = (getattr(report, "workflow", None) or {})
    n_buttons = len(wf.get("buttons", []) or [])
    n_events = len(wf.get("event_handlers", []) or [])
    if n_buttons or n_events:
        L.append(f"- **Operational entry points**: {n_buttons} button(s), "
                 f"{n_events} event handler(s). See Workflow Guide for the walkthrough.")
    else:
        L.append(f"- **No buttons or event handlers detected** — workbook is "
                 f"formula-driven only. Users interact via cell entries directly.")

    # Risks
    r = report.risk_indicators
    risks_short = []
    if r.very_hidden_sheets:
        risks_short.append(f"{len(r.very_hidden_sheets)} veryHidden sheet(s)")
    if r.hidden_sheets:
        risks_short.append(f"{len(r.hidden_sheets)} hidden")
    if r.cells_with_errors:
        risks_short.append(f"{len(r.cells_with_errors)} formula-error cell(s)")
    if r.external_workbook_references:
        risks_short.append(f"{len(r.external_workbook_references)} external link(s)")
    if risks_short:
        L.append(f"- **Risk flags**: " + "; ".join(risks_short) + ".")

    # Sub-score table (compact)
    L.append("")
    L.append("**Complexity sub-scores:**")
    L.append("")
    L.append("| Sub-score | Value | Bar |")
    L.append("|---|---|---|")
    sub = c.sub_scores
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        L.append(f"| {key.replace('_', ' ')} | {val}/20 | `{_sub_score_bar(val)}` |")
    L.append("")
    return L


# ---------------------------------------------------------------------------
# Reference Appendix sub-builders (re-use existing _section_* with light
# wrappers to relabel anchors as 8.x).
# ---------------------------------------------------------------------------

def _section_appendix_intro(report) -> list:
    L: list = []
    L.append("## Reference Appendix")
    L.append("")
    L.append("_Full data tables and indices. Every catalog the audit produces "
             "lives below — for technical readers verifying findings or following "
             "up on a Top Impact entry._")
    L.append("")
    return L


# VBA classification summary table (for D6 — replaces the unreadable diagram
# with a class-level table at the top, then a bounded mini-diagram).
def _section_vba_classification_summary(report) -> list:
    L: list = []
    L.append("### VBA classification overview")
    L.append("")
    classifications = report.vba_classifications or []
    if not classifications:
        L.append("_No VBA modules._")
        L.append("")
        return L
    by_class: dict = defaultdict(list)
    for c in classifications:
        by_class[c.inferred_type].append(c)
    # Build module -> LOC lookup
    loc_by_name = {vm.name: vm.line_count for vm in report.vba_modules}

    L.append("| Type | Count | Total LOC | Sample modules |")
    L.append("|---|---|---|---|")
    known_order = ["data-loader", "transformer", "report-writer",
                   "ui-handler", "dead-suspected", "mixed"]
    seen: set = set()
    for cls in known_order:
        items = by_class.get(cls, [])
        if not items:
            continue
        seen.add(cls)
        total_loc = sum(loc_by_name.get(c.module_name, 0) for c in items)
        sample = ", ".join(f"`{c.module_name}`" for c in items[:3])
        if len(items) > 3:
            sample += f" (+{len(items)-3} more)"
        L.append(f"| `{cls}` | {len(items)} | {total_loc:,} | {sample} |")
    for cls, items in sorted(by_class.items()):
        if cls in seen:
            continue
        total_loc = sum(loc_by_name.get(c.module_name, 0) for c in items)
        sample = ", ".join(f"`{c.module_name}`" for c in items[:3])
        if len(items) > 3:
            sample += f" (+{len(items)-3} more)"
        L.append(f"| `{cls}` | {len(items)} | {total_loc:,} | {sample} |")
    L.append("")
    L.append("**Classification mini-diagram (top 2 categories, ≤ 15 nodes):**")
    L.append("")
    L.append("```mermaid")
    L.append(_build_vba_classification_mermaid(report))
    L.append("```")
    L.append("")
    return L


# ---------------------------------------------------------------------------
# Top-level Markdown
# ---------------------------------------------------------------------------

# Round-3 PYRAMID structure (D2): cover → exec → workflow → dataflow →
# top-impact → vba-walkthrough → domain → appendix → glossary → methodology.
# Each header still appears exactly once. Appendix sub-headers are H3 under
# the single H2 "Reference Appendix".
_TOP_LEVEL_HEADERS_FOR_TOC = [
    "Executive Summary",
    "Workflow Guide",
    "Data Flow Story",
    "Top Impact Findings",
    "VBA Module Walkthrough",
    "Domain-Specific Findings",
    "Reference Appendix",
    "Glossary",
    "Methodology",
]


def _exec_summary_lines(report) -> list:
    """Build the executive summary block. Returns markdown lines."""
    return _section_executive_summary_round3(report)


def render_markdown(report) -> str:
    """Round-3 pyramid layout — selective narrative > exhaustive dump."""
    return _assemble_markdown(report)


def _section_pillar_impact_diagram_only(report) -> list:
    """Standalone Pillar impact diagram (H3 under appendix)."""
    L: list = []
    L.append(f"### Pillar impact diagram")
    L.append("")
    L.append(f"_Top-{_DIAG3_MAX_PILLARS} pillar cells and the sheets they cascade into._")
    L.append("")
    L.append("```mermaid")
    L.append(_build_pillar_impact_mermaid(report))
    L.append("```")
    L.append("")
    return L


def _domain_findings_present(report) -> bool:
    matches = (report.domain_template_matches or [])
    matches = [m for m in matches if m.confidence in ("high", "medium")]
    return bool(matches)


def _assemble_markdown(report) -> str:
    """Round-3 pyramid assembly.

    1. Cover (filename, audit version, sanitize banner)
    2. Executive Summary (manager-readable)
    3. Workflow Guide (NEW)
    4. Data Flow Story (NEW)
    5. Top Impact Findings (Top-N each)
    6. VBA Module Walkthrough (NEW)
    7. Domain-Specific Findings (when domain detected)
    8. Reference Appendix (H3 sub-sections — full data)
    9. Glossary (NEW)
    10. Methodology
    """
    L: list = []
    m = report.meta
    from . import __version__ as _pkg_version
    # 1. Cover
    headline_complexity = report.complexity.total
    L.append(f"# Audit report — `{m.file_name}` (audit v{_pkg_version})")
    L.append("")

    if report.sanitized:
        L.append("> 🔒 **SANITIZED MODE** — no cell values in this report. "
                 "Formulas, structure, smells, and VBA source are preserved; "
                 "every non-formula cell value has been replaced with `<redacted>`.")
        L.append(">")
    L.append(f"> Headline: complexity **{headline_complexity}/100**, "
             f"**{len(report.pillars)}** pillar cell(s), "
             f"**{len(report.smells)}** smell finding(s).")
    L.append(">")
    L.append("> Tier 1 audit. Pure static analysis — no AI, no Excel, no macro execution.")
    L.append("> Same input always produces the same output. Findings ranked, not interpreted.")
    L.append("")

    # Build TOC dynamically (skip Domain-Specific Findings if not detected)
    toc_headers = list(_TOP_LEVEL_HEADERS_FOR_TOC)
    if not _domain_findings_present(report):
        toc_headers = [h for h in toc_headers if h != "Domain-Specific Findings"]

    # 2. Executive Summary
    L.extend(_exec_summary_lines(report))

    # TOC (right after exec)
    L.append("## Table of Contents")
    L.append("")
    for header in toc_headers:
        anchor = _slugify_anchor(header)
        L.append(f"- [{header}](#{anchor})")
    L.append("")

    # 3. Workflow Guide
    L.extend(_section_workflow_guide(report))

    # 4. Data Flow Story
    L.extend(_section_data_flow_story(report))

    # 5. Top Impact Findings
    L.extend(_section_top_impact_findings(report))

    # 6. VBA Module Walkthrough
    L.extend(_section_vba_walkthrough(report))

    # 7. Domain-Specific Findings (only when present)
    if _domain_findings_present(report):
        L.extend(_section_domain_findings(report))

    # 8. Reference Appendix (full data tables, H3-level)
    L.extend(_section_appendix_intro(report))

    # 8.1 Pillar table (full)
    L.append("### 8.1 Full pillar table")
    L.append("")
    pillar_rows = _section_pillars(report, top_n=None, with_drilldown=True)
    # _section_pillars emits its own ## H2 header; replace with the H3 we want.
    # Cleaner: build full table here by extracting rows.
    if not report.pillars:
        L.append("_No cells reach the pillar threshold._")
        L.append("")
    else:
        L.append("_Full deduped pillar list — see Top Impact §5 for the Top-5 view._")
        L.append("")
        L.append("| Rank | Cell / range | Value | Label | Members | Fan-in | Affected sheets | Kind | Narrative |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        for i, p in enumerate(report.pillars, start=1):
            sheets_str = ", ".join(f"`{s}`" for s in p.affected_sheets[:5])
            if len(p.affected_sheets) > 5:
                sheets_str += f" (+{len(p.affected_sheets) - 5} more)"
            L.append(
                f"| {i} | `{_md_escape(p.location)}` | {_format_pillar_value_cell(p)} | "
                f"{_format_pillar_label_cell(p)} | {p.member_count} | {p.fan_in} | "
                f"{sheets_str} | {p.pillar_kind} | {_md_escape(p.narrative)} |"
            )
        L.append("")

    # 8.2 Smell catalog (full)
    L.append("### 8.2 Full smells catalog (Hermans 2015)")
    L.append("")
    L.append(f"_{len(report.smells)} smell finding(s) across "
             f"{len({s.smell_type for s in report.smells})} smell type(s)._")
    L.append("")
    by_type: dict = defaultdict(list)
    for s in report.smells:
        by_type[s.smell_type].append(s)
    for st in SMELL_TYPES:
        items = by_type.get(st, [])
        L.append(f"#### `{st}` — {len(items)} finding(s)")
        L.append("")
        if not items:
            L.append("_No findings above threshold._")
            L.append("")
            continue
        L.append("| Location | Metric | Severity | Confidence | Evidence |")
        L.append("|---|---|---|---|---|")
        for s in items[:20]:
            L.append(
                f"| `{_md_escape(s.location)}` | {s.metric:g} | {s.severity} | "
                f"{s.confidence} | {_md_escape(s.evidence)} |"
            )
        if len(items) > 20:
            L.append("")
            L.append(f"_({len(items)-20} more findings of this type — see `audit.json`.)_")
        L.append("")

    # 8.3 Sheets table
    L.append("### 8.3 Sheets")
    L.append("")
    L.append("| Sheet | State | Rows | Cols | Non-empty | Formula | Max ref | CF | DV |")
    L.append("|---|---|---|---|---|---|---|---|---|")
    for s in report.sheets:
        L.append(
            f"| `{_md_escape(s.name)}` | {s.state} | {s.rows_used} | {s.cols_used} | "
            f"{s.cells_nonempty} | {s.cells_formula} | {s.max_ref or '—'} | "
            f"{s.conditional_formatting_count} | {s.data_validation_count} |"
        )
    L.append("")

    # 8.4 Named ranges
    L.append("### 8.4 Named ranges")
    L.append("")
    if not report.named_ranges:
        L.append("_No named ranges defined._")
    else:
        L.append("| Name | Scope | Reference |")
        L.append("|---|---|---|")
        for nr in report.named_ranges:
            L.append(f"| `{_md_escape(nr.name)}` | `{_md_escape(nr.scope)}` | `{_md_escape(nr.ref)}` |")
    L.append("")

    # 8.5 Magic-number index
    L.append("### 8.5 Magic-number index (top 20)")
    L.append("")
    if not report.magic_numbers:
        L.append("_No non-trivial numeric literals found._")
    else:
        L.append("| Value | Count | First location | Source | Sample context |")
        L.append("|---|---|---|---|---|")
        for mn in report.magic_numbers:
            L.append(
                f"| `{_md_escape(mn.value)}` | {mn.occurrence_count} | "
                f"`{_md_escape(mn.first_location)}` | {mn.location_kind} | "
                f"`{_md_escape(mn.sample_context)}` |"
            )
    L.append("")

    # 8.6 Risk indicators (full)
    r = report.risk_indicators
    L.append("### 8.6 Risk indicators")
    L.append("")
    L.append("| Indicator | Value |")
    L.append("|---|---|")
    L.append(f"| Hidden sheets | {len(r.hidden_sheets)} (`{', '.join(r.hidden_sheets) if r.hidden_sheets else '—'}`) |")
    L.append(f"| Very-hidden sheets | {len(r.very_hidden_sheets)} (`{', '.join(r.very_hidden_sheets) if r.very_hidden_sheets else '—'}`) |")
    L.append(f"| Cross-sheet referencing formulas | {r.cross_sheet_reference_count} |")
    L.append(f"| Cells with formula errors (cached) | {len(r.cells_with_errors)} |")
    L.append(f"| External workbook reference patterns | {len(r.external_workbook_references)} |")
    L.append(f"| Circular reference suspects | {len(r.circular_reference_suspects)} |")
    L.append(f"| Parse errors logged | {len(r.parse_errors)} |")
    L.append("")
    if r.cells_with_errors:
        L.append("**Formula error cells (first 20):**")
        L.append("")
        L.append("| Sheet | Ref | Error |")
        L.append("|---|---|---|")
        for ec in r.cells_with_errors[:20]:
            L.append(f"| `{_md_escape(ec['sheet'])}` | `{ec['ref']}` | `{ec['error_token']}` |")
        L.append("")
    if r.external_workbook_references:
        L.append("**External workbook references:**")
        L.append("")
        for ext in r.external_workbook_references[:20]:
            L.append(f"- `{_md_escape(ext)}`")
        L.append("")
    if r.circular_reference_suspects:
        L.append("**Circular reference suspects (first 20):**")
        L.append("")
        for cs in r.circular_reference_suspects[:20]:
            L.append(f"- `{_md_escape(cs)}`")
        L.append("")

    # 8.7 Complexity score breakdown
    c = report.complexity
    L.append("### 8.7 Complexity score breakdown")
    L.append("")
    L.append(f"**Total: {c.total} / 100**")
    L.append("")
    L.append("| Sub-score | Value | Bar | Rationale |")
    L.append("|---|---|---|---|")
    sub = c.sub_scores
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        bar = _sub_score_bar(val)
        L.append(f"| {key.replace('_',' ')} | {val}/20 | `{bar}` | {_md_escape(c.rationale[key])} |")
    L.append("")

    # 8.8 VBA classification + full table
    L.extend(_section_vba_classification_summary(report))
    # Full VBA modules table
    L.append("#### VBA modules — full table")
    L.append("")
    if not report.vba_modules:
        L.append("_No VBA modules found._")
    else:
        cls_by_name = {c.module_name: c for c in (report.vba_classifications or [])}
        L.append("| Module | Type | LOC | #Sub | #Func | Inferred type | Confidence | Reads | Writes | Ext calls | OnErrorResumeNext |")
        L.append("|---|---|---|---|---|---|---|---|---|---|---|")
        for vm in report.vba_modules:
            n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
            n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
            cls = cls_by_name.get(vm.name)
            inferred = cls.inferred_type if cls else "—"
            cls_conf = cls.confidence if cls else "—"
            reads = ", ".join(f"`{s}`" for s in (cls.reads_sheets if cls else [])) or "—"
            writes = ", ".join(f"`{s}`" for s in (cls.writes_sheets if cls else [])) or "—"
            ext_call = ("yes" if (cls and cls.external_calls) else "no") if cls else "—"
            L.append(
                f"| `{_md_escape(vm.name)}` | {vm.type} | {vm.line_count} | "
                f"{n_sub} | {n_func} | **{inferred}** | {cls_conf} | "
                f"{_md_escape(reads)} | {_md_escape(writes)} | {ext_call} | "
                f"{'yes' if vm.has_on_error_resume_next else 'no'} |"
            )
    L.append("")

    # 8.9 Pillar impact diagram (kept as evidence)
    L.extend(_section_pillar_impact_diagram_only(report))

    # 8.10 File metadata + basic stats
    L.append("### 8.10 File metadata")
    L.append("")
    L.append("| Field | Value |")
    L.append("|---|---|")
    L.append(f"| File name | `{m.file_name}` |")
    L.append(f"| File size | {m.file_size_bytes:,} bytes ({m.file_size_bytes/1024:.1f} KB) |")
    L.append(f"| SHA-256 | `{m.sha256}` |")
    if report.sanitized:
        L.append(f"| Sanitize mode | **active** (cell values redacted) |")
    L.append("")

    # 8.11 Basic statistics (the original table)
    b = report.basic_stats
    L.append("### 8.11 Basic statistics")
    L.append("")
    L.append("| Metric | Value |")
    L.append("|---|---|")
    L.append(f"| Sheet count (total) | {b.sheet_count} |")
    L.append(f"| Sheet count visible / hidden / veryHidden | "
             f"{b.sheet_count_visible} / {b.sheet_count_hidden} / {b.sheet_count_very_hidden} |")
    L.append(f"| Non-empty cells | {b.cell_count_nonempty:,} |")
    L.append(f"| Formula cells | {b.cell_count_formula:,} |")
    L.append(f"| Unique non-formula values | {b.cell_count_unique_values:,} |")
    L.append(f"| Named ranges | {b.named_range_count} |")
    L.append(f"| Conditional formatting rules | {b.conditional_formatting_count} |")
    L.append(f"| Data validation rules | {b.data_validation_count} |")
    L.append(f"| VBA modules | {b.vba_module_count} |")
    L.append(f"| VBA total lines | {b.vba_total_lines:,} |")
    L.append(f"| Cell-level parse errors (logged + skipped) | {b.parse_errors_count} |")
    L.append("")

    # 9. Glossary
    L.extend(_section_glossary(report))

    # 10. Methodology
    L.extend(_section_methodology(report))

    return "\n".join(L) + "\n"


# ---------------------------------------------------------------------------
# HTML render
# ---------------------------------------------------------------------------

_CSS = """
:root { color-scheme: light; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC",
               "Hiragino Sans GB", "Microsoft YaHei", Arial, sans-serif;
  font-size: 14px;
  line-height: 1.55;
  color: #1f2328;
  background: #ffffff;
  max-width: 1100px;
  margin: 24px auto;
  padding: 0 24px;
}
h1 { font-size: 22px; border-bottom: 1px solid #d0d7de; padding-bottom: 8px; }
h2 { font-size: 18px; margin-top: 28px; padding-top: 6px; border-bottom: 1px solid #eaecef; }
h3 { font-size: 15px; margin-top: 18px; }
.audit-section { page-break-inside: avoid; }
.audit-section + .audit-section { page-break-before: auto; }
@media print {
  body { max-width: 100%; margin: 0; padding: 12px; font-size: 11pt; }
  h2 { page-break-before: auto; page-break-after: avoid; }
  .toc { page-break-after: always; }
  pre, table { page-break-inside: avoid; }
  .mermaid { page-break-inside: avoid; }
}
table { border-collapse: collapse; margin: 8px 0 16px 0; font-size: 13px; }
th, td { border: 1px solid #d0d7de; padding: 5px 9px; text-align: left; vertical-align: top; }
th { background: #f6f8fa; font-weight: 600; }
code, pre { font-family: SFMono-Regular, Consolas, "Liberation Mono", monospace; }
code { background: #f6f8fa; padding: 1px 4px; border-radius: 3px; font-size: 0.92em; }
pre { background: #f6f8fa; padding: 12px; border-radius: 6px; overflow-x: auto; }
pre code { background: transparent; padding: 0; }
blockquote {
  border-left: 4px solid #d0d7de;
  margin: 12px 0;
  padding: 4px 16px;
  color: #57606a;
  background: #f6f8fa;
}
.exec-summary {
  background: #fff8e1;
  border: 1px solid #f7c948;
  border-radius: 8px;
  padding: 16px 20px;
  margin: 16px 0;
}
.exec-summary h2 { border-bottom: none; margin-top: 0; }
.toc {
  background: #f6f8fa;
  border: 1px solid #d0d7de;
  border-radius: 6px;
  padding: 12px 20px;
}
.toc ul { margin: 4px 0; padding-left: 20px; }
.callout-card {
  background: #e3f2fd;
  border-left: 4px solid #1976d2;
  padding: 8px 14px;
  margin: 8px 0;
  border-radius: 4px;
}
.callout-anomaly { border-left-color: #d84315; background: #fbe9e7; }
.callout-pillar { border-left-color: #6a1b9a; background: #f3e5f5; }
.bar { font-family: monospace; letter-spacing: -1px; }
.banner-sanitize {
  background: #fff3cd;
  border: 1px solid #f0ad4e;
  border-radius: 6px;
  padding: 10px 14px;
  margin: 12px 0;
  font-weight: 600;
}
.mermaid { background: #fafbfc; padding: 12px; border: 1px solid #eaecef; border-radius: 6px; overflow-x: auto; }
"""


def _h(s: Any) -> str:
    """HTML-escape any value as text content."""
    return _html.escape(str(s), quote=True)


def _html_table(headers: list, rows: list) -> str:
    out = ['<table><thead><tr>']
    for h in headers:
        out.append(f"<th>{_h(h)}</th>")
    out.append("</tr></thead><tbody>")
    for r in rows:
        out.append("<tr>")
        for c in r:
            out.append(f"<td>{c}</td>")  # cells may already be HTML-escaped
        out.append("</tr>")
    out.append("</tbody></table>")
    return "".join(out)


def _html_code(s: Any) -> str:
    return f"<code>{_h(s)}</code>"


def _html_section_open(title: str) -> str:
    anchor = _slugify_anchor(title)
    return f'<section class="audit-section" id="{anchor}"><h2>{_h(title)}</h2>'


def _html_section_close() -> str:
    return "</section>"


def _html_exec_summary(report) -> str:
    c = report.complexity
    domain = getattr(report, "domain_hint", None)
    findings = _headline_findings(report)

    parts: list = []
    parts.append('<section class="audit-section exec-summary" id="executive-summary">')
    parts.append('<h2>Executive Summary</h2>')
    parts.append(f"<p><strong>Complexity</strong>: {c.total} / 100</p>")

    sub = c.sub_scores
    rows = []
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        bar_html = f'<span class="bar">{_h(_sub_score_bar(val))}</span>'
        rows.append([
            _h(key.replace("_", " ")),
            _h(f"{val}/20"),
            bar_html,
        ])
    parts.append(_html_table(["Sub-score", "Value", "Bar"], rows))

    parts.append("<p><strong>Top findings:</strong></p>")
    if findings:
        cards: list = []
        for f in findings:
            cls = "callout-card"
            if f.startswith("**Pillar"):
                cls = "callout-card callout-pillar"
            elif f.startswith("**Anomaly"):
                cls = "callout-card callout-anomaly"
            cards.append(f'<div class="{cls}">{_inline_md_to_html(f)}</div>')
        parts.append("".join(cards))
    else:
        parts.append("<p><em>No salient findings.</em></p>")

    if domain is not None:
        if domain.domain == "unknown":
            parts.append(
                "<p><strong>Detected domain</strong>: <em>unknown</em> — domain "
                "not auto-detected; analyze manually.</p>"
            )
        else:
            kw_str = ", ".join(domain.matched_keywords)
            parts.append(
                f"<p><strong>Detected domain</strong>: <code>{_h(domain.domain)}</code> "
                f"<em>(confidence: {_h(domain.confidence)} — matched: {_h(kw_str)})</em></p>"
            )
    parts.append("</section>")
    return "".join(parts)


_RE_INLINE_BOLD = re.compile(r"\*\*(.+?)\*\*")
_RE_INLINE_CODE = re.compile(r"`([^`]+?)`")
_RE_INLINE_EMPH = re.compile(r"_([^_]+?)_")


def _inline_md_to_html(s: str) -> str:
    """Render the small subset of inline markdown we use (`code`, **bold**, _emph_) to HTML.
    The input is trusted (built by us). We escape the residue for safety."""
    # Escape first, then re-render the markup tokens
    escaped = _h(s)
    # Re-replace patterns based on the *escaped* string — markdown tokens stay
    # intact because **, `, _ are not HTML-special characters.
    out = _RE_INLINE_CODE.sub(r"<code>\1</code>", escaped)
    out = _RE_INLINE_BOLD.sub(r"<strong>\1</strong>", out)
    out = _RE_INLINE_EMPH.sub(r"<em>\1</em>", out)
    return out


def _html_table_section(report, title: str, headers: list, rows: list, intro_html: str = "") -> str:
    out = [_html_section_open(title)]
    if intro_html:
        out.append(intro_html)
    out.append(_html_table(headers, rows))
    out.append(_html_section_close())
    return "".join(out)


def _html_file_metadata(report) -> str:
    m = report.meta
    rows = [
        ["File name", _html_code(m.file_name)],
        ["File size", _h(f"{m.file_size_bytes:,} bytes ({m.file_size_bytes/1024:.1f} KB)")],
        ["SHA-256", _html_code(m.sha256)],
    ]
    if report.sanitized:
        rows.append(["Sanitize mode", "<strong>active</strong> (cell values redacted)"])
    return _html_table_section(report, "File metadata", ["Field", "Value"], rows)


def _html_basic_stats(report) -> str:
    b = report.basic_stats
    rows = [
        ["Sheet count (total)", _h(b.sheet_count)],
        ["Sheet count visible / hidden / veryHidden",
         _h(f"{b.sheet_count_visible} / {b.sheet_count_hidden} / {b.sheet_count_very_hidden}")],
        ["Non-empty cells", _h(f"{b.cell_count_nonempty:,}")],
        ["Formula cells", _h(f"{b.cell_count_formula:,}")],
        ["Unique non-formula values", _h(f"{b.cell_count_unique_values:,}")],
        ["Named ranges", _h(b.named_range_count)],
        ["Conditional formatting rules", _h(b.conditional_formatting_count)],
        ["Data validation rules", _h(b.data_validation_count)],
        ["VBA modules", _h(b.vba_module_count)],
        ["VBA total lines", _h(f"{b.vba_total_lines:,}")],
        ["Cell-level parse errors (logged + skipped)", _h(b.parse_errors_count)],
    ]
    return _html_table_section(report, "Basic statistics", ["Metric", "Value"], rows)


def _html_sheets(report) -> str:
    rows = []
    for s in report.sheets:
        rows.append([
            _html_code(s.name),
            _h(s.state),
            _h(s.rows_used),
            _h(s.cols_used),
            _h(s.cells_nonempty),
            _h(s.cells_formula),
            _h(s.max_ref or "—"),
            _h(s.conditional_formatting_count),
            _h(s.data_validation_count),
        ])
    return _html_table_section(
        report, "Sheets",
        ["Sheet", "State", "Rows", "Cols", "Non-empty", "Formula", "Max ref", "CF", "DV"],
        rows,
    )


def _html_named_ranges(report) -> str:
    if not report.named_ranges:
        return (_html_section_open("Named ranges")
                + "<p><em>No named ranges defined.</em></p>"
                + _html_section_close())
    rows = [
        [_html_code(nr.name), _html_code(nr.scope), _html_code(nr.ref)]
        for nr in report.named_ranges
    ]
    return _html_table_section(report, "Named ranges", ["Name", "Scope", "Reference"], rows)


def _html_complexity(report) -> str:
    c = report.complexity
    sub = c.sub_scores
    rows = []
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        rows.append([
            _h(key.replace("_", " ")),
            _h(f"{val}/20"),
            f'<span class="bar">{_h(_sub_score_bar(val))}</span>',
            _h(c.rationale[key]),
        ])
    intro = f"<p><strong>Total</strong>: {_h(c.total)} / 100</p>"
    return _html_table_section(
        report, "Complexity score",
        ["Sub-score", "Value", "Bar", "Rationale"],
        rows, intro_html=intro,
    )


def _html_pillars(report) -> str:
    out = [_html_section_open("Pillar cells — systemic single-points-of-impact")]
    out.append('<p><em>Cells with the highest fan-in (most-referenced). Modifying any '
               "of these cascades through many formulas; treat them as critical change "
               "points. Equivalent column cells (same fan-in, same affected sheets) are grouped.</em></p>")
    if not report.pillars:
        thr = report.methodology["logic_depth_thresholds"]["pillar-fanin-min"]
        out.append(f"<p><em>No cells reach the pillar threshold (fan-in &ge; {_h(thr)}).</em></p>")
        out.append(_html_section_close())
        return "".join(out)

    rows = []
    for i, p in enumerate(report.pillars, start=1):
        sheets_str = ", ".join(_html_code(s) for s in p.affected_sheets[:5])
        if len(p.affected_sheets) > 5:
            sheets_str += f" (+{len(p.affected_sheets) - 5} more)"
        rows.append([
            _h(i), _html_code(p.location), _h(p.member_count), _h(p.fan_in),
            sheets_str, _h(p.pillar_kind), "high", _h(p.narrative),
        ])
    out.append(_html_table(
        ["Rank", "Cell / range", "Members", "Fan-in", "Affected sheets",
         "Kind", "Confidence", "Narrative"],
        rows,
    ))

    out.append("<h3>Top-5 pillar drilldown — sample dependents</h3>")
    out.append("<ul>")
    for p in report.pillars[:5]:
        if p.member_count > 1:
            heading = (f"<code>{_h(p.location)}</code> ({p.fan_in} dependents per cell × "
                       f"{p.member_count} cells)")
        else:
            heading = f"<code>{_h(p.location)}</code> ({p.fan_in} dependents)"
        cls = "callout-card callout-pillar"
        out.append(f'<li><div class="{cls}"><strong>{heading}</strong><ul>')
        for d in p.sample_dependents:
            out.append(f"<li><code>{_h(d)}</code></li>")
        if p.member_count > 1 and p.member_refs:
            preview = ", ".join(f"<code>{_h(r)}</code>" for r in p.member_refs[:3])
            extra = f" (+{len(p.member_refs) - 3} more)" if len(p.member_refs) > 3 else ""
            out.append(f"<li><em>Group members:</em> {preview}{extra}</li>")
        out.append("</ul></div></li>")
    out.append("</ul>")
    out.append(_html_section_close())
    return "".join(out)


def _html_anomalies(report) -> str:
    out = [_html_section_open("Magic-number anomalies — outliers within formula clusters")]
    out.append("<p><em>Inside groups of cells that share the same formula shape, this "
               "section flags positions where a small minority uses a different numeric "
               "constant. These are exactly the kind of \"why is this row's discount "
               "different?\" findings that often indicate a missed update or a deliberate "
               "(but undocumented) carve-out.</em></p>")
    if not report.anomalies:
        out.append("<p><em>No magic-number anomalies detected. Either no large duplicated-formula "
                   "clusters were found, or every cluster's numeric constants are perfectly consistent.</em></p>")
        out.append(_html_section_close())
        return "".join(out)
    rows = []
    for i, a in enumerate(report.anomalies, start=1):
        locs = ", ".join(_html_code(loc) for loc in a.outlier_locations[:5])
        if len(a.outlier_locations) > 5:
            locs += f" (+{len(a.outlier_locations) - 5} more)"
        rows.append([
            _h(i),
            _html_code(a.cluster_pattern_sample),
            _h(a.cluster_size),
            f"{_html_code(a.mode_value)} ({a.mode_count}/{a.cluster_size})",
            f"{_html_code(a.outlier_value)} ({a.outlier_count}/{a.cluster_size})",
            locs,
            _h(a.confidence),
            _h(a.narrative),
        ])
    out.append(_html_table(
        ["#", "Cluster sample", "Cluster size", "Mode value",
         "Outlier value", "Outlier locations", "Confidence", "Narrative"],
        rows,
    ))
    out.append(_html_section_close())
    return "".join(out)


def _html_smells(report) -> str:
    out = [_html_section_open("Smells catalog (Hermans 2015)")]
    out.append(f"<p><em>{_h(len(report.smells))} smell findings across "
               f"{_h(len({s.smell_type for s in report.smells}))} smell types.</em></p>")
    by_type: dict = defaultdict(list)
    for s in report.smells:
        by_type[s.smell_type].append(s)
    for st in SMELL_TYPES:
        items = by_type.get(st, [])
        out.append(f"<h3><code>{_h(st)}</code> &mdash; {_h(len(items))} findings</h3>")
        if not items:
            out.append("<p><em>No findings above threshold.</em></p>")
            continue
        rows = []
        for s in items[:20]:
            rows.append([
                _html_code(s.location), _h(f"{s.metric:g}"),
                _h(s.severity), _h(s.confidence), _h(s.evidence),
            ])
        out.append(_html_table(
            ["Location", "Metric", "Severity", "Confidence", "Evidence"],
            rows,
        ))
        if len(items) > 20:
            out.append(f"<p><em>({len(items)-20} more findings of this type — see "
                       "<code>audit.json</code>.)</em></p>")
    out.append(_html_section_close())
    return "".join(out)


def _html_magic_index(report) -> str:
    if not report.magic_numbers:
        return (_html_section_open("Magic-number index (top 20)")
                + "<p><em>No non-trivial numeric literals found.</em></p>"
                + _html_section_close())
    rows = [
        [
            _html_code(mn.value),
            _h(mn.occurrence_count),
            _html_code(mn.first_location),
            _h(mn.location_kind),
            _html_code(mn.sample_context),
        ]
        for mn in report.magic_numbers
    ]
    return _html_table_section(
        report, "Magic-number index (top 20)",
        ["Value", "Count", "First location", "Source", "Sample context"],
        rows,
    )


def _html_vba(report) -> str:
    out = [_html_section_open("VBA modules + classification")]
    if not report.vba_modules:
        out.append("<p><em>No VBA modules found.</em></p>")
        out.append(_html_section_close())
        return "".join(out)
    cls_by_name = {c.module_name: c for c in (report.vba_classifications or [])}

    rows = []
    for vm in report.vba_modules:
        n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
        n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
        cls = cls_by_name.get(vm.name)
        inferred = cls.inferred_type if cls else "—"
        cls_conf = cls.confidence if cls else "—"
        reads = ", ".join(_html_code(s) for s in (cls.reads_sheets if cls else [])) or "—"
        writes = ", ".join(_html_code(s) for s in (cls.writes_sheets if cls else [])) or "—"
        ext_call = ("yes" if (cls and cls.external_calls) else "no") if cls else "—"
        rows.append([
            _html_code(vm.name), _h(vm.type), _h(vm.line_count),
            _h(n_sub), _h(n_func),
            f"<strong>{_h(inferred)}</strong>",
            _h(cls_conf), reads, writes, _h(ext_call),
            _h("yes" if vm.has_on_error_resume_next else "no"),
        ])
    out.append(_html_table(
        ["Module", "Type", "LOC", "#Sub", "#Func", "Inferred type",
         "Confidence", "Reads", "Writes", "Ext calls", "OnErrorResumeNext"],
        rows,
    ))

    out.append("<h3>VBA module details — external keywords &amp; range literals</h3>")
    rows2 = []
    for vm in report.vba_modules:
        ext = ", ".join(vm.external_keywords) or "—"
        ranges = len(vm.range_literals)
        cls = cls_by_name.get(vm.name)
        rationale = cls.rationale if cls else "—"
        rows2.append([
            _html_code(vm.name), _h(ext), _h(ranges), _h(rationale),
        ])
    out.append(_html_table(
        ["Module", "External keywords", "Range literals (uniq)", "Classifier rationale"],
        rows2,
    ))
    out.append(_html_section_close())
    return "".join(out)


def _html_risks(report) -> str:
    out = [_html_section_open("Risk indicators")]
    r = report.risk_indicators
    rows = [
        ["Hidden sheets",
         _h(f"{len(r.hidden_sheets)} ({', '.join(r.hidden_sheets) if r.hidden_sheets else '—'})")],
        ["Very-hidden sheets",
         _h(f"{len(r.very_hidden_sheets)} ({', '.join(r.very_hidden_sheets) if r.very_hidden_sheets else '—'})")],
        ["Cross-sheet referencing formulas", _h(r.cross_sheet_reference_count)],
        ["Cells with formula errors (cached)", _h(len(r.cells_with_errors))],
        ["External workbook reference patterns", _h(len(r.external_workbook_references))],
        ["Circular reference suspects", _h(len(r.circular_reference_suspects))],
        ["Parse errors logged", _h(len(r.parse_errors))],
    ]
    out.append(_html_table(["Indicator", "Value"], rows))

    if r.cells_with_errors:
        out.append("<h3>Formula error cells (first 20)</h3>")
        rows = [
            [_html_code(ec["sheet"]), _html_code(ec["ref"]), _html_code(ec["error_token"])]
            for ec in r.cells_with_errors[:20]
        ]
        out.append(_html_table(["Sheet", "Ref", "Error"], rows))
    if r.external_workbook_references:
        out.append("<h3>External workbook references</h3><ul>")
        for ext in r.external_workbook_references[:20]:
            out.append(f"<li>{_html_code(ext)}</li>")
        out.append("</ul>")
    if r.circular_reference_suspects:
        out.append("<h3>Circular reference suspects (first 20)</h3><ul>")
        for cs in r.circular_reference_suspects[:20]:
            out.append(f"<li>{_html_code(cs)}</li>")
        out.append("</ul>")
    out.append(_html_section_close())
    return "".join(out)


def _html_diagrams(report) -> str:
    out: list = []
    for title, intro, builder in [
        ("Sheet data flow diagram",
         "Cross-sheet formula references. Edge label = number of formulas with "
         "cross-sheet refs from source -> target sheet. Yellow = hidden, red = "
         "veryHidden, default = visible.",
         _build_sheet_dataflow_mermaid),
        ("VBA classification overview diagram",
         "Modules grouped under their inferred type.",
         _build_vba_classification_mermaid),
        ("Pillar impact diagram",
         f"Top-{_DIAG3_MAX_PILLARS} pillar cells and the sheets they cascade into.",
         _build_pillar_impact_mermaid),
    ]:
        out.append(_html_section_open(title))
        out.append(f"<p><em>{_h(intro)}</em></p>")
        # Mermaid source — escaped because Mermaid block is read as text by the JS
        # but inserted as inner-HTML of a <pre class="mermaid">. We must escape.
        out.append(f'<pre class="mermaid">{_h(builder(report))}</pre>')
        out.append(_html_section_close())
    return "".join(out)


def _html_methodology(report) -> str:
    out = [_html_section_open("Methodology")]
    libs = report.methodology["library_versions"]
    out.append(f"<ul><li>Engine: <code>openpyxl</code> {_h(libs['openpyxl'])} "
               f"(cells/structure), <code>oletools.olevba</code> {_h(libs['oletools'])} "
               f"(VBA), <code>formulas</code> {_h(libs['formulas'])} (tokenizer only — "
               "no evaluator).</li>")
    thr = report.methodology["smell_thresholds"]
    out.append(
        f"<li>Smell thresholds: multiple-references &ge; {_h(thr['multiple-references'])}; "
        f"long-calculation-chain depth &ge; {_h(thr['long-calculation-chain'])}; "
        f"conditional-complexity nesting &ge; {_h(thr['conditional-complexity'])}; "
        f"multiple-operations &ge; {_h(thr['multiple-operations'])}; "
        f"duplicated-formulas pattern frequency &ge; {_h(thr['duplicated-formulas'])}.</li>"
    )
    ld = report.methodology["logic_depth_thresholds"]
    out.append(
        f"<li>Logic-depth thresholds: pillar fan-in &ge; {_h(ld['pillar-fanin-min'])} "
        f"(top {_h(ld['pillar-top-n'])} after dedupe); anomaly cluster size &ge; "
        f"{_h(ld['anomaly-cluster-min-size'])}, outlier fraction &le; "
        f"{_h(ld['anomaly-outlier-fraction'])}.</li>"
    )
    out.append("<li>Pillar dedupe: cells in the same column with identical fan-in and "
               "identical affected-sheet set are collapsed into one column-block entry.</li>")
    out.append("<li>VBA classifier categories: " +
               ", ".join(_html_code(c) for c in report.methodology["vba_classifier_categories"]) +
               ".</li>")
    out.append("<li>Confidence semantics: <code>high</code> = exact / deterministic count; "
               "<code>medium</code> = tokenizer-based with well-defined rules; "
               "<code>low</code> = statistical inference or analysis skipped due to scale.</li>")
    out.append("<li>Domain hint detector: pure keyword matching (case-insensitive, "
               "word-boundary). No LLM, no inference.</li>")
    out.append("<li>Trivial numbers excluded from magic-number index: " +
               ", ".join(_html_code(n) for n in sorted(TRIVIAL_NUMBERS)) + ".</li>")
    out.append("<li>Reliability contract: same input &rarr; byte-identical "
               "<code>audit.md</code>, <code>audit.json</code>, and <code>audit.html</code>. "
               "All outputs use sorted dict keys and lexicographic list ordering. "
               "No timestamps. No CDN fetch at audit time.</li>")
    if report.sanitized:
        out.append("<li><strong>Sanitize mode active</strong>: every non-formula cell value "
                   "has been replaced with <code>&lt;redacted&gt;</code> before any analysis ran. "
                   "Formulas, VBA source, smells, pillars, anomalies, and structural counts "
                   "remain accurate. The SHA-256 above is of the original file.</li>")
    out.append("</ul>")
    out.append("<hr><p><em>End of report.</em></p>")
    out.append(_html_section_close())
    return "".join(out)


def _html_toc(report) -> str:
    """Round-3 TOC: only the top-level pyramid headers, in order."""
    headers = [h for h in _TOP_LEVEL_HEADERS_FOR_TOC]
    if not _domain_findings_present(report):
        headers = [h for h in headers if h != "Domain-Specific Findings"]
    out: list = ['<nav class="toc audit-section" id="toc"><h2>Table of Contents</h2><ul>']
    for header in headers:
        anchor = _slugify_anchor(header)
        out.append(f'<li><a href="#{anchor}">{_h(header)}</a></li>')
    out.append("</ul></nav>")
    return "".join(out)


# ---------------------------------------------------------------------------
# Round-3 HTML section builders for the new prose-heavy sections.
# These reuse the markdown builders and convert the result to HTML using
# the small inline-MD subset we already render in callouts.
# ---------------------------------------------------------------------------

def _md_lines_to_html(lines: list) -> str:
    """Convert a list of markdown lines (as produced by section builders)
    into HTML. Handles: H2/H3/H4 headers, paragraphs, blockquotes, bullet
    lists, fenced code (mermaid), and tables with pipe syntax. Inline
    `code`, **bold**, _emph_ are converted in place. We're NOT a full
    markdown engine — just enough for this section's output.
    """
    out: list = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        # H2
        if line.startswith("## "):
            title = line[3:].strip()
            anchor = _slugify_anchor(title)
            out.append(f'<section class="audit-section" id="{anchor}"><h2>{_h(title)}</h2>')
            i += 1
            continue
        if line.startswith("### "):
            out.append(f"<h3>{_h(line[4:].strip())}</h3>")
            i += 1
            continue
        if line.startswith("#### "):
            out.append(f"<h4>{_h(line[5:].strip())}</h4>")
            i += 1
            continue
        # Mermaid fenced code
        if line.strip() == "```mermaid":
            j = i + 1
            buf = []
            while j < n and lines[j].strip() != "```":
                buf.append(lines[j])
                j += 1
            out.append(f'<pre class="mermaid">{_h(chr(10).join(buf))}</pre>')
            i = j + 1
            continue
        # Generic fenced code
        if line.strip().startswith("```"):
            j = i + 1
            buf = []
            while j < n and lines[j].strip() != "```":
                buf.append(lines[j])
                j += 1
            out.append(f'<pre><code>{_h(chr(10).join(buf))}</code></pre>')
            i = j + 1
            continue
        # Table: line starts with `|` AND next line is `|---|...|`
        if line.lstrip().startswith("|") and i + 1 < n and re.match(r"^\s*\|[\s\-|]+\|\s*$", lines[i+1]):
            # parse header + separator + body until non-pipe line
            header_cells = [c.strip() for c in line.strip().strip("|").split("|")]
            i += 2
            rows: list = []
            while i < n and lines[i].lstrip().startswith("|"):
                row_cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(row_cells)
                i += 1
            html_rows = []
            for r in rows:
                html_rows.append([_inline_md_to_html(c) for c in r])
            out.append(_html_table(header_cells, html_rows))
            continue
        # HTML comment (LLM-AUGMENT marker) — pass through verbatim
        if line.lstrip().startswith("<!--"):
            out.append(line)
            i += 1
            continue
        # Bullet list — accumulate
        if line.lstrip().startswith("- "):
            ul = ["<ul>"]
            while i < n and lines[i].lstrip().startswith("- "):
                stripped = lines[i].lstrip()[2:]
                ul.append(f"<li>{_inline_md_to_html(stripped)}</li>")
                i += 1
            ul.append("</ul>")
            out.append("".join(ul))
            continue
        # Blank line — close any open paragraph (no-op here since we don't
        # buffer paragraphs; just emit a small spacer)
        if not line.strip():
            i += 1
            continue
        # Default: paragraph
        # Collect consecutive non-blank, non-special lines into one paragraph
        buf: list = [line]
        i += 1
        while i < n:
            nxt = lines[i]
            if (not nxt.strip() or nxt.startswith(("#", "|", "- ", "```"))
                    or nxt.lstrip().startswith("<!--")):
                break
            buf.append(nxt)
            i += 1
        para_text = " ".join(b.strip() for b in buf)
        out.append(f"<p>{_inline_md_to_html(para_text)}</p>")
    # Close any unclosed sections (every <section> opens with </section> not emitted)
    # We count open sections and close them at the end:
    open_sections = sum(1 for s in out if s.startswith('<section')) - sum(
        1 for s in out if s == "</section>")
    for _ in range(open_sections):
        out.append("</section>")
    return "".join(out)


def _html_workflow_guide(report) -> str:
    return _md_lines_to_html(_section_workflow_guide(report))


def _html_data_flow_story(report) -> str:
    return _md_lines_to_html(_section_data_flow_story(report))


def _html_top_impact_findings(report) -> str:
    return _md_lines_to_html(_section_top_impact_findings(report))


def _html_vba_walkthrough(report) -> str:
    return _md_lines_to_html(_section_vba_walkthrough(report))


def _html_domain_findings(report) -> str:
    return _md_lines_to_html(_section_domain_findings(report))


def _html_glossary(report) -> str:
    return _md_lines_to_html(_section_glossary(report))


def _html_appendix(report) -> str:
    """Render the Reference Appendix using the existing tableized helpers
    (so the appendix retains rich-HTML tables — preferred over MD-derived)."""
    parts: list = []
    parts.append('<section class="audit-section" id="reference-appendix"><h2>Reference Appendix</h2>')
    parts.append("<p><em>Full data tables and indices. Every catalog the audit "
                 "produces lives below — for technical readers verifying findings "
                 "or following up on a Top Impact entry.</em></p>")
    # 8.1 Pillar table (full) — reuse _html_pillars but rename heading
    parts.append("<h3>8.1 Full pillar table</h3>")
    parts.append(_html_pillars_table_only(report))
    # 8.2 Smells catalog
    parts.append("<h3>8.2 Full smells catalog (Hermans 2015)</h3>")
    parts.append(_html_smells_inner(report))
    # 8.3 Sheets
    parts.append("<h3>8.3 Sheets</h3>")
    parts.append(_html_sheets_inner(report))
    # 8.4 Named ranges
    parts.append("<h3>8.4 Named ranges</h3>")
    parts.append(_html_named_ranges_inner(report))
    # 8.5 Magic-number index
    parts.append("<h3>8.5 Magic-number index (top 20)</h3>")
    parts.append(_html_magic_index_inner(report))
    # 8.6 Risk indicators
    parts.append("<h3>8.6 Risk indicators</h3>")
    parts.append(_html_risks_inner(report))
    # 8.7 Complexity breakdown
    parts.append("<h3>8.7 Complexity score breakdown</h3>")
    parts.append(_html_complexity_inner(report))
    # 8.8 VBA classification + table
    parts.append("<h3>8.8 VBA classification overview</h3>")
    parts.append(_html_vba_summary_inner(report))
    parts.append("<h4>VBA modules — full table</h4>")
    parts.append(_html_vba_inner(report))
    # 8.9 Pillar impact diagram
    parts.append("<h3>8.9 Pillar impact diagram</h3>")
    parts.append(f"<p><em>Top-{_DIAG3_MAX_PILLARS} pillar cells and the sheets they cascade into.</em></p>")
    parts.append(f'<pre class="mermaid">{_h(_build_pillar_impact_mermaid(report))}</pre>')
    # 8.10 File metadata
    parts.append("<h3>8.10 File metadata</h3>")
    parts.append(_html_file_metadata_inner(report))
    # 8.11 Basic statistics
    parts.append("<h3>8.11 Basic statistics</h3>")
    parts.append(_html_basic_stats_inner(report))
    parts.append("</section>")
    return "".join(parts)


# Inner table helpers — strip the section wrapper so the appendix can place
# them under H3 sub-headers rather than separate H2 sections.
def _html_pillars_table_only(report) -> str:
    if not report.pillars:
        thr = report.methodology["logic_depth_thresholds"]["pillar-fanin-min"]
        return f"<p><em>No cells reach the pillar threshold (fan-in &ge; {_h(thr)}).</em></p>"
    rows = []
    for i, p in enumerate(report.pillars, start=1):
        sheets_str = ", ".join(_html_code(s) for s in p.affected_sheets[:5])
        if len(p.affected_sheets) > 5:
            sheets_str += f" (+{len(p.affected_sheets) - 5} more)"
        # Value cell HTML
        v = (getattr(p, "value", "") or "").strip()
        val_html = "<em>(empty)</em>" if not v else f"<code>{_h(v[:30])}</code>"
        # Label cell HTML
        label_parts = []
        if getattr(p, "named_range", ""):
            label_parts.append(f"named range <code>{_h(p.named_range)}</code>")
        if getattr(p, "row_header", ""):
            label_parts.append(f"row label <code>{_h(p.row_header[:25])}</code>")
        if getattr(p, "col_header", "") and p.col_header != getattr(p, "row_header", ""):
            label_parts.append(f"col header <code>{_h(p.col_header[:25])}</code>")
        label_html = "; ".join(label_parts) if label_parts else "—"
        rows.append([
            _h(i), _html_code(p.location), val_html, label_html,
            _h(p.member_count), _h(p.fan_in),
            sheets_str, _h(p.pillar_kind), _h(p.narrative),
        ])
    return _html_table(
        ["Rank", "Cell / range", "Value", "Label", "Members", "Fan-in",
         "Affected sheets", "Kind", "Narrative"],
        rows,
    )


def _html_smells_inner(report) -> str:
    out = [f"<p><em>{_h(len(report.smells))} smell finding(s) across "
           f"{_h(len({s.smell_type for s in report.smells}))} smell type(s).</em></p>"]
    by_type: dict = defaultdict(list)
    for s in report.smells:
        by_type[s.smell_type].append(s)
    for st in SMELL_TYPES:
        items = by_type.get(st, [])
        out.append(f"<h4><code>{_h(st)}</code> &mdash; {_h(len(items))} finding(s)</h4>")
        if not items:
            out.append("<p><em>No findings above threshold.</em></p>")
            continue
        rows = []
        for s in items[:20]:
            rows.append([
                _html_code(s.location), _h(f"{s.metric:g}"),
                _h(s.severity), _h(s.confidence), _h(s.evidence),
            ])
        out.append(_html_table(
            ["Location", "Metric", "Severity", "Confidence", "Evidence"], rows))
        if len(items) > 20:
            out.append(f"<p><em>({len(items)-20} more findings of this type — see "
                       "<code>audit.json</code>.)</em></p>")
    return "".join(out)


def _html_sheets_inner(report) -> str:
    rows = []
    for s in report.sheets:
        rows.append([
            _html_code(s.name), _h(s.state), _h(s.rows_used), _h(s.cols_used),
            _h(s.cells_nonempty), _h(s.cells_formula),
            _h(s.max_ref or "—"),
            _h(s.conditional_formatting_count), _h(s.data_validation_count),
        ])
    return _html_table(
        ["Sheet", "State", "Rows", "Cols", "Non-empty", "Formula", "Max ref", "CF", "DV"],
        rows,
    )


def _html_named_ranges_inner(report) -> str:
    if not report.named_ranges:
        return "<p><em>No named ranges defined.</em></p>"
    rows = [
        [_html_code(nr.name), _html_code(nr.scope), _html_code(nr.ref)]
        for nr in report.named_ranges
    ]
    return _html_table(["Name", "Scope", "Reference"], rows)


def _html_magic_index_inner(report) -> str:
    if not report.magic_numbers:
        return "<p><em>No non-trivial numeric literals found.</em></p>"
    rows = [
        [
            _html_code(mn.value), _h(mn.occurrence_count),
            _html_code(mn.first_location), _h(mn.location_kind),
            _html_code(mn.sample_context),
        ]
        for mn in report.magic_numbers
    ]
    return _html_table(
        ["Value", "Count", "First location", "Source", "Sample context"], rows)


def _html_risks_inner(report) -> str:
    out: list = []
    r = report.risk_indicators
    rows = [
        ["Hidden sheets",
         _h(f"{len(r.hidden_sheets)} ({', '.join(r.hidden_sheets) if r.hidden_sheets else '—'})")],
        ["Very-hidden sheets",
         _h(f"{len(r.very_hidden_sheets)} ({', '.join(r.very_hidden_sheets) if r.very_hidden_sheets else '—'})")],
        ["Cross-sheet referencing formulas", _h(r.cross_sheet_reference_count)],
        ["Cells with formula errors (cached)", _h(len(r.cells_with_errors))],
        ["External workbook reference patterns", _h(len(r.external_workbook_references))],
        ["Circular reference suspects", _h(len(r.circular_reference_suspects))],
        ["Parse errors logged", _h(len(r.parse_errors))],
    ]
    out.append(_html_table(["Indicator", "Value"], rows))
    if r.cells_with_errors:
        out.append("<h4>Formula error cells (first 20)</h4>")
        rows = [
            [_html_code(ec["sheet"]), _html_code(ec["ref"]), _html_code(ec["error_token"])]
            for ec in r.cells_with_errors[:20]
        ]
        out.append(_html_table(["Sheet", "Ref", "Error"], rows))
    if r.external_workbook_references:
        out.append("<h4>External workbook references</h4><ul>")
        for ext in r.external_workbook_references[:20]:
            out.append(f"<li>{_html_code(ext)}</li>")
        out.append("</ul>")
    if r.circular_reference_suspects:
        out.append("<h4>Circular reference suspects (first 20)</h4><ul>")
        for cs in r.circular_reference_suspects[:20]:
            out.append(f"<li>{_html_code(cs)}</li>")
        out.append("</ul>")
    return "".join(out)


def _html_complexity_inner(report) -> str:
    c = report.complexity
    sub = c.sub_scores
    rows = []
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        rows.append([
            _h(key.replace("_", " ")),
            _h(f"{val}/20"),
            f'<span class="bar">{_h(_sub_score_bar(val))}</span>',
            _h(c.rationale[key]),
        ])
    return (f"<p><strong>Total</strong>: {_h(c.total)} / 100</p>"
            + _html_table(["Sub-score", "Value", "Bar", "Rationale"], rows))


def _html_vba_summary_inner(report) -> str:
    """Render VBA classification summary + the bounded mermaid diagram."""
    classifications = report.vba_classifications or []
    if not classifications:
        return "<p><em>No VBA modules.</em></p>"
    by_class: dict = defaultdict(list)
    for c in classifications:
        by_class[c.inferred_type].append(c)
    loc_by_name = {vm.name: vm.line_count for vm in report.vba_modules}
    known_order = ["data-loader", "transformer", "report-writer",
                   "ui-handler", "dead-suspected", "mixed"]
    rows = []
    seen: set = set()
    for cls in known_order:
        items = by_class.get(cls, [])
        if not items:
            continue
        seen.add(cls)
        total_loc = sum(loc_by_name.get(c.module_name, 0) for c in items)
        sample = ", ".join(_html_code(c.module_name) for c in items[:3])
        if len(items) > 3:
            sample += f" (+{len(items)-3} more)"
        rows.append([_html_code(cls), _h(len(items)), _h(f"{total_loc:,}"), sample])
    for cls, items in sorted(by_class.items()):
        if cls in seen:
            continue
        total_loc = sum(loc_by_name.get(c.module_name, 0) for c in items)
        sample = ", ".join(_html_code(c.module_name) for c in items[:3])
        if len(items) > 3:
            sample += f" (+{len(items)-3} more)"
        rows.append([_html_code(cls), _h(len(items)), _h(f"{total_loc:,}"), sample])
    out = [_html_table(["Type", "Count", "Total LOC", "Sample modules"], rows)]
    out.append("<h4>Classification mini-diagram (top 2 categories, &le; 15 nodes)</h4>")
    out.append(f'<pre class="mermaid">{_h(_build_vba_classification_mermaid(report))}</pre>')
    return "".join(out)


def _html_vba_inner(report) -> str:
    if not report.vba_modules:
        return "<p><em>No VBA modules found.</em></p>"
    cls_by_name = {c.module_name: c for c in (report.vba_classifications or [])}
    rows = []
    for vm in report.vba_modules:
        n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
        n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
        cls = cls_by_name.get(vm.name)
        inferred = cls.inferred_type if cls else "—"
        cls_conf = cls.confidence if cls else "—"
        reads = ", ".join(_html_code(s) for s in (cls.reads_sheets if cls else [])) or "—"
        writes = ", ".join(_html_code(s) for s in (cls.writes_sheets if cls else [])) or "—"
        ext_call = ("yes" if (cls and cls.external_calls) else "no") if cls else "—"
        rows.append([
            _html_code(vm.name), _h(vm.type), _h(vm.line_count),
            _h(n_sub), _h(n_func),
            f"<strong>{_h(inferred)}</strong>",
            _h(cls_conf), reads, writes, _h(ext_call),
            _h("yes" if vm.has_on_error_resume_next else "no"),
        ])
    return _html_table(
        ["Module", "Type", "LOC", "#Sub", "#Func", "Inferred type",
         "Confidence", "Reads", "Writes", "Ext calls", "OnErrorResumeNext"],
        rows,
    )


def _html_file_metadata_inner(report) -> str:
    m = report.meta
    rows = [
        ["File name", _html_code(m.file_name)],
        ["File size", _h(f"{m.file_size_bytes:,} bytes ({m.file_size_bytes/1024:.1f} KB)")],
        ["SHA-256", _html_code(m.sha256)],
    ]
    if report.sanitized:
        rows.append(["Sanitize mode", "<strong>active</strong> (cell values redacted)"])
    return _html_table(["Field", "Value"], rows)


def _html_basic_stats_inner(report) -> str:
    b = report.basic_stats
    rows = [
        ["Sheet count (total)", _h(b.sheet_count)],
        ["Sheet count visible / hidden / veryHidden",
         _h(f"{b.sheet_count_visible} / {b.sheet_count_hidden} / {b.sheet_count_very_hidden}")],
        ["Non-empty cells", _h(f"{b.cell_count_nonempty:,}")],
        ["Formula cells", _h(f"{b.cell_count_formula:,}")],
        ["Unique non-formula values", _h(f"{b.cell_count_unique_values:,}")],
        ["Named ranges", _h(b.named_range_count)],
        ["Conditional formatting rules", _h(b.conditional_formatting_count)],
        ["Data validation rules", _h(b.data_validation_count)],
        ["VBA modules", _h(b.vba_module_count)],
        ["VBA total lines", _h(f"{b.vba_total_lines:,}")],
        ["Cell-level parse errors (logged + skipped)", _h(b.parse_errors_count)],
    ]
    return _html_table(["Metric", "Value"], rows)


_MERMAID_CDN = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"


def render_html(report, mermaid_inline: bool = False, mermaid_inline_source: str = "") -> str:
    """Render the audit as a styled HTML page (round-3 pyramid layout)."""
    from . import __version__ as _pkg_version
    m = report.meta

    parts: list = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="en"><head>')
    parts.append('<meta charset="utf-8">')
    parts.append('<meta name="viewport" content="width=device-width,initial-scale=1">')
    parts.append(f"<title>Audit: {_h(m.file_name)}</title>")
    parts.append(f"<style>{_CSS}</style>")
    if mermaid_inline:
        safe_js = mermaid_inline_source.replace("</script>", "<\\/script>")
        parts.append(f"<script>{safe_js}</script>")
    else:
        parts.append(f'<script src="{_MERMAID_CDN}"></script>')
    parts.append("<script>"
                 "if (window.mermaid) { mermaid.initialize({startOnLoad:true, securityLevel:'loose'}); }"
                 "</script>")
    parts.append("</head><body>")

    parts.append(f"<h1>Audit report &mdash; <code>{_h(m.file_name)}</code> "
                 f"(audit v{_h(_pkg_version)})</h1>")

    if report.sanitized:
        parts.append('<div class="banner-sanitize">&#128274; SANITIZED MODE — no cell '
                     "values in this report. Formulas, structure, smells, and VBA source "
                     "are preserved; every non-formula cell value has been replaced with "
                     "<code>&lt;redacted&gt;</code>.</div>")

    headline_complexity = report.complexity.total
    parts.append(f'<blockquote>Headline: complexity <strong>{headline_complexity}/100</strong>, '
                 f'<strong>{len(report.pillars)}</strong> pillar cell(s), '
                 f'<strong>{len(report.smells)}</strong> smell finding(s).<br>'
                 'Tier 1 audit. Pure static analysis — no AI, no Excel, '
                 "no macro execution. Same input always produces the same output. "
                 "Findings ranked, not interpreted.</blockquote>")

    # Pyramid order:
    parts.append(_html_exec_summary(report))
    parts.append(_html_toc(report))
    parts.append(_html_workflow_guide(report))
    parts.append(_html_data_flow_story(report))
    parts.append(_html_top_impact_findings(report))
    parts.append(_html_vba_walkthrough(report))
    if _domain_findings_present(report):
        parts.append(_html_domain_findings(report))
    parts.append(_html_appendix(report))
    parts.append(_html_glossary(report))
    parts.append(_html_methodology(report))

    parts.append("</body></html>")
    return "".join(parts) + "\n"
