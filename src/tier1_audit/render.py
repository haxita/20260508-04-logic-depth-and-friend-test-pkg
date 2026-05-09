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

from .i18n import (
    DEFAULT_LANG,
    SUPPORTED_LANGS,
    md_table_separator,
    split_pipe_columns,
    t,
)
from .smells import SMELL_TYPES, TRIVIAL_NUMBERS

# Field names to omit from JSON serialization (huge or noisy or internal-only)
_EXCLUDED_FIELDS = frozenset({"source_text", "_sheet_edges"})


# ---------------------------------------------------------------------------
# i18n helpers — re-derive narratives from structured fields per language
# ---------------------------------------------------------------------------

def _truncate_label_render(s: str, max_len: int = 30) -> str:
    """Truncate a label/value to max_len, appending an ellipsis if cut.
    Mirrors pillars._truncate_label so render-time re-derivation matches."""
    if not s:
        return s
    if len(s) <= max_len:
        return s
    return s[: max_len - 1].rstrip() + "…"


def _pillar_inline_phrase(p, lang: str) -> str:
    """Build the inline ' (value `X`, named range `Y`, label `Z`)' clause
    using i18n labels. Returns '' when nothing useful is known (graceful).

    Intentionally kept structurally identical to pillars._value_label_phrase
    so the same Pillar object produces an equivalent inline phrase in any
    language — only the prefix words ('value', 'named range', 'label') change.
    """
    parts: list = []
    val = getattr(p, "value", "") or ""
    nr = getattr(p, "named_range", "") or ""
    rh = getattr(p, "row_header", "") or ""
    ch = getattr(p, "col_header", "") or ""
    if val:
        # English keeps the verbatim "value" prefix from the master narrative;
        # other langs translate just the prefix word.
        if lang == DEFAULT_LANG:
            parts.append(f"value `{_truncate_label_render(val, 30)}`")
        else:
            # Re-use the row/col label prefix style — but for value we need
            # an explicit translation. For DE/ZH we hardcode to keep this
            # function self-contained (the prefix words are stable enough
            # to live in the narrative.pillar.* keys we already added).
            value_word = {"de": "Wert", "zh": "值"}.get(lang, "value")
            parts.append(f"{value_word} `{_truncate_label_render(val, 30)}`")
    if nr:
        prefix = t("pillar.label_named_range_prefix", lang)
        parts.append(f"{prefix} `{nr}`")
    label = rh or ch
    if label and label != val:
        prefix = t("pillar.label_row_prefix", lang)
        parts.append(f"{prefix} `{_truncate_label_render(label, 30)}`")
    if not parts:
        return ""
    return " (" + ", ".join(parts) + ")"


def _pillar_kind_word(p, lang: str) -> str:
    """Translate pillar_kind into the user-visible kind word."""
    pk = getattr(p, "pillar_kind", "")
    if pk == "column-block":
        return t("narrative.pillar.kind_column_block", lang)
    if pk == "formula-relay":
        return t("narrative.pillar.kind_formula_relay", lang)
    if pk == "constant-input":
        return t("narrative.pillar.kind_value_anchor", lang)
    return t("narrative.pillar.kind_default", lang)


def _pillar_scope_phrase(p, lang: str) -> str:
    """Render the 'sheet/sheets …' scope clause."""
    sheets = list(p.affected_sheets or [])
    n = len(sheets)
    if n == 1:
        return t("narrative.pillar.scope_one", lang, sheet=sheets[0])
    if n <= 3:
        joined = ", ".join(f"`{s}`" for s in sheets)
        return t("narrative.pillar.scope_few", lang, sheets=joined)
    listed = ", ".join(f"`{s}`" for s in sheets[:3])
    return t("narrative.pillar.scope_many", lang, n_sheets=n, listed=listed)


def _pillar_risk_phrase(fan_in: int, lang: str) -> str:
    if fan_in >= 100:
        return t("narrative.pillar.risk_high", lang)
    if fan_in >= 50:
        return t("narrative.pillar.risk_med", lang)
    return t("narrative.pillar.risk_low", lang)


def _render_pillar_narrative(p, lang: str) -> str:
    """Re-derive the pillar narrative for `lang`.

    For English we return the master narrative stored on the Pillar (so the
    JSON output is byte-identical with pre-i18n behavior on EN). For DE / ZH
    we rebuild from structured fields using i18n templates.
    """
    if lang == DEFAULT_LANG:
        return getattr(p, "narrative", "") or ""
    inline = _pillar_inline_phrase(p, lang)
    scope = _pillar_scope_phrase(p, lang)
    risk = _pillar_risk_phrase(p.fan_in, lang)
    kind_word = _pillar_kind_word(p, lang)
    if p.member_count > 1:
        return t(
            "narrative.pillar.group", lang,
            member_count=p.member_count,
            inline=inline,
            fan_in=p.fan_in,
            scope=scope,
            risk=risk,
        )
    return t(
        "narrative.pillar.single", lang,
        location=p.location,
        inline=inline,
        kind_word=kind_word,
        fan_in=p.fan_in,
        scope=scope,
        risk=risk,
    )


def _render_anomaly_narrative(a, lang: str) -> str:
    """Re-derive the anomaly narrative for `lang`."""
    if lang == DEFAULT_LANG:
        return getattr(a, "narrative", "") or ""
    locs = list(getattr(a, "outlier_locations", []) or [])
    if len(locs) == 1:
        cell_phrase = t("narrative.anomaly.cell_phrase_one", lang,
                        location=locs[0])
    else:
        first = locs[0] if locs else ""
        cell_phrase = t("narrative.anomaly.cell_phrase_many", lang,
                        n=a.outlier_count, first=first)
    mode_share = int(round(a.mode_count / max(a.cluster_size, 1) * 100))
    # deviation: same approximation as anomalies._make_narrative — use
    # numeric difference between mode and outlier when both numeric.
    try:
        dev = abs(float(a.outlier_value) - float(a.mode_value))
        dev_str = f"{dev:g}"
    except (TypeError, ValueError):
        dev_str = "n/a"
    return t(
        "narrative.anomaly.text", lang,
        cluster_size=a.cluster_size,
        sample_formula=a.cluster_pattern_sample,
        position=a.position_index + 1,
        mode_display=a.mode_value,
        mode_count=a.mode_count,
        mode_share=mode_share,
        cell_phrase=cell_phrase,
        outlier_display=a.outlier_value,
        dev=dev_str,
    )


def _render_vba_role_inference(role_key: str, lang: str) -> str:
    """Translate the VbaNarrative.role_inference free-text back via key.

    The narrate module emits role_inference as English prose. We pattern-match
    on the canonical English text to find the i18n key. If no match, we keep
    the raw English (better than `[[missing:…]]` for unexpected roles)."""
    if lang == DEFAULT_LANG:
        return role_key
    role_key_norm = role_key.strip()
    role_lookup = {
        "isolated module — not reached from any button or event handler (possibly dead)":
            "narrative.vba.role_isolated",
        "leaf helper — invoked by other modules but invokes nothing further":
            "narrative.vba.role_leaf_helper",
        "system entry point — invokes other modules and is not invoked back":
            "narrative.vba.role_entry_point",
        "large multi-purpose module — likely the workbook's main logic block":
            "narrative.vba.role_large",
        "data ingest — reads external sources or workbook ranges into VBA structures":
            "narrative.vba.role_data_loader",
        "compute step — reads inputs, derives outputs, writes back to sheets":
            "narrative.vba.role_transformer",
        "output writer — populates sheet ranges with computed values":
            "narrative.vba.role_report_writer",
        "UI event handler — fires when the user opens the workbook or interacts with a sheet":
            "narrative.vba.role_ui_handler",
        "near-empty shell — typical for default class/sheet stubs with no real code":
            "narrative.vba.role_dead_shell",
        "mixed responsibilities — no single structural signal dominates":
            "narrative.vba.role_mixed",
    }
    key = role_lookup.get(role_key_norm)
    if key is None:
        return role_key
    return t(key, lang)


def _render_vba_narrative(narr, lang: str) -> str:
    """Re-derive the VBA narrative paragraph from VbaNarrative fields.

    For English: we keep the existing pre-built narr.narrative for byte-
    identical output. For DE/ZH: rebuild via i18n templates so structurally
    equivalent prose appears in the target language.
    """
    if lang == DEFAULT_LANG:
        return getattr(narr, "narrative", "") or ""
    lines: list = []
    role_translated = _render_vba_role_inference(narr.role_inference, lang)
    lines.append(t("narrative.vba.role_inference", lang, role=role_translated))

    parts: list = []
    if narr.reads_sheets:
        sheets = ", ".join(f"`{s}`" for s in narr.reads_sheets[:5])
        more = (t("common.more_suffix", lang, n=len(narr.reads_sheets) - 5)
                if len(narr.reads_sheets) > 5 else "")
        parts.append(t("narrative.vba.reads", lang, sheets=sheets, more=more))
    if narr.writes_sheets:
        sheets = ", ".join(f"`{s}`" for s in narr.writes_sheets[:5])
        more = (t("common.more_suffix", lang, n=len(narr.writes_sheets) - 5)
                if len(narr.writes_sheets) > 5 else "")
        parts.append(t("narrative.vba.writes", lang, sheets=sheets, more=more))
    if narr.callees:
        c = ", ".join(f"`{x}`" for x in narr.callees[:3])
        more = (t("common.more_suffix", lang, n=len(narr.callees) - 3)
                if len(narr.callees) > 3 else "")
        parts.append(t("narrative.vba.calls_into", lang, callees=c, more=more))
    if narr.callers and not narr.callees:
        c = ", ".join(f"`{x}`" for x in narr.callers[:3])
        parts.append(t("narrative.vba.invoked_by", lang, callers=c))
    if narr.nested_loops_max >= 2:
        parts.append(t("narrative.vba.nested_loops", lang,
                       n=narr.nested_loops_max))
    if not parts:
        parts.append(t("narrative.vba.no_io", lang))
    lines.append(t("narrative.vba.what_does", lang, parts="; ".join(parts)))

    # Notable patterns
    if narr.notable_patterns:
        lines.append(t("narrative.vba.notable_patterns", lang))
        for bp in narr.notable_patterns:
            lines.append(f"- {_translate_notable_pattern(bp, narr, lang)}")

    # Call relationships
    rel_parts: list = []
    if narr.callers:
        c = ", ".join(f"`{x}`" for x in narr.callers[:3])
        more = (t("common.more_suffix", lang, n=len(narr.callers) - 3)
                if len(narr.callers) > 3 else "")
        rel_parts.append(t("narrative.vba.called_by", lang, callers=c, more=more))
    else:
        rel_parts.append(t("narrative.vba.no_callers", lang))
    if narr.callees:
        c = ", ".join(f"`{x}`" for x in narr.callees[:3])
        more = (t("common.more_suffix", lang, n=len(narr.callees) - 3)
                if len(narr.callees) > 3 else "")
        rel_parts.append(t("narrative.vba.calls_out", lang, callees=c, more=more))
    else:
        rel_parts.append(t("narrative.vba.no_callees", lang))
    lines.append(t("narrative.vba.relations", lang, parts="; ".join(rel_parts)))

    return "\n".join(lines)


_RE_OER_LINES = re.compile(
    r"Contains `On Error Resume Next` at line\(s\) ([0-9, ]+)"
)
_RE_NESTED_LOOPS = re.compile(
    r"Contains (\d+) levels of nested loops"
)
_RE_EXTERNAL_KW = re.compile(
    r"Uses external/COM API keyword\(s\): (.+?)$"
)


def _translate_notable_pattern(bullet: str, narr, lang: str) -> str:
    """Translate the small set of notable-pattern bullets emitted by
    vba_narrate._build_notable_patterns. Pattern-match the English shape and
    rebuild via i18n. Falls back to the raw English bullet if no match."""
    if lang == DEFAULT_LANG:
        return bullet
    m = _RE_OER_LINES.search(bullet)
    if m:
        return t("narrative.vba.oer_with_lines", lang, lines=m.group(1).strip())
    if "Contains `On Error Resume Next` —" in bullet:
        return t("narrative.vba.oer_no_lines", lang)
    m = _RE_NESTED_LOOPS.search(bullet)
    if m:
        return t("narrative.vba.nested_loops_pattern", lang, n=m.group(1))
    if bullet.startswith("Uses external/COM API keyword"):
        # Pull keywords out
        # Form: "Uses external/COM API keyword(s): `foo`, `bar` (+N more)."
        rest = bullet[len("Uses external/COM API keyword(s):"):].strip().rstrip(".")
        # Try to split (+N more) tail
        more_match = re.search(r"\(\+(\d+) more\)\s*$", rest)
        more = ""
        if more_match:
            more = t("common.more_suffix", lang, n=more_match.group(1))
            rest = rest[: more_match.start()].rstrip()
        return t("narrative.vba.external_keywords", lang,
                 keywords=rest, more=more)
    return bullet


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


def render_json(report, lang: str = DEFAULT_LANG) -> str:
    """Render the audit report as JSON.

    For English (default), the output is byte-identical to the pre-i18n
    behaviour. For DE / ZH, narrative text fields (pillar.narrative,
    anomaly.narrative, vba_narrative.narrative, vba_narrative.role_inference,
    vba_narrative.notable_patterns) are re-derived in the target language;
    all structural fields stay verbatim.
    """
    data = _to_jsonable(report)
    if lang != DEFAULT_LANG:
        data = _translate_jsonable(data, report, lang)
    return json.dumps(data, sort_keys=True, indent=2, ensure_ascii=False) + "\n"


def _translate_jsonable(data: dict, report, lang: str) -> dict:
    """Walk the JSON-able dict and translate the small set of narrative
    fields that contain prose. All other fields stay verbatim — sheet names,
    cell values, formula text, and numeric counts are user data.

    Implementation note: we walk by key paths the audit pipeline uses, which
    is small and stable. We do NOT attempt a generic "translate every string
    field" walk — that would risk mangling user data.
    """
    # Pillars: re-derive p.narrative from structured fields
    pillars_in = data.get("pillars") or []
    if pillars_in and report.pillars:
        for i, p_dict in enumerate(pillars_in):
            if i < len(report.pillars):
                p_dict["narrative"] = _render_pillar_narrative(report.pillars[i], lang)

    # Anomalies
    anomalies_in = data.get("anomalies") or []
    if anomalies_in and report.anomalies:
        for i, a_dict in enumerate(anomalies_in):
            if i < len(report.anomalies):
                a_dict["narrative"] = _render_anomaly_narrative(report.anomalies[i], lang)

    # VBA narratives
    narratives_in = data.get("vba_narratives") or []
    if narratives_in and report.vba_narratives:
        for i, n_dict in enumerate(narratives_in):
            if i < len(report.vba_narratives):
                narr = report.vba_narratives[i]
                n_dict["narrative"] = _render_vba_narrative(narr, lang)
                n_dict["role_inference"] = _render_vba_role_inference(
                    narr.role_inference, lang
                )
                # notable_patterns: translate each bullet
                n_dict["notable_patterns"] = [
                    _translate_notable_pattern(bp, narr, lang)
                    for bp in (narr.notable_patterns or [])
                ]
    return data


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

def _headline_findings(report, lang: str = DEFAULT_LANG) -> list:
    """Return up to 3 single-line headline findings: Top pillar, top anomaly
    or top smell, plus one risk indicator. Strings only — no markup beyond
    inline backticks."""
    out: list = []
    if report.pillars:
        p = report.pillars[0]
        if p.member_count > 1:
            out.append(t("exec_summary.headline_pillar_group", lang,
                         location=p.location,
                         member_count=p.member_count,
                         fan_in=p.fan_in))
        else:
            out.append(t("exec_summary.headline_pillar_single", lang,
                         location=p.location,
                         fan_in=p.fan_in,
                         affected_sheet_count=p.affected_sheet_count))

    if report.anomalies:
        a = report.anomalies[0]
        out.append(t("exec_summary.headline_anomaly", lang,
                     position=a.position_index + 1,
                     cluster_size=a.cluster_size,
                     mode_value=a.mode_value,
                     outlier_value=a.outlier_value,
                     outlier_location=a.outlier_locations[0],
                     confidence=a.confidence))
    elif report.smells:
        top_smell = max(report.smells, key=lambda s: s.metric)
        out.append(t("exec_summary.top_smell", lang,
                     smell_type=top_smell.smell_type,
                     location=top_smell.location,
                     metric=f"{top_smell.metric:g}",
                     severity=top_smell.severity))

    r = report.risk_indicators
    risks = []
    if r.very_hidden_sheets:
        risks.append(t("exec_summary.headline_risk_very_hidden", lang,
                       n=len(r.very_hidden_sheets)))
    if r.hidden_sheets:
        risks.append(t("exec_summary.headline_risk_hidden", lang,
                       n=len(r.hidden_sheets)))
    if r.cells_with_errors:
        risks.append(t("exec_summary.headline_risk_errors", lang,
                       n=len(r.cells_with_errors)))
    if r.external_workbook_references:
        risks.append(t("exec_summary.headline_risk_external", lang,
                       n=len(r.external_workbook_references)))
    if r.circular_reference_suspects:
        risks.append(t("exec_summary.headline_risk_circular", lang,
                       n=len(r.circular_reference_suspects)))
    if risks:
        out.append(t("exec_summary.headline_risks", lang, risks="; ".join(risks)))

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

def _md_table_header_lines(columns_key: str, lang: str) -> list:
    """Build (header_row, separator_row) markdown lines from a `|`-joined
    catalog value. Used for table-header keys."""
    cols = split_pipe_columns(t(columns_key, lang))
    return [
        "| " + " | ".join(cols) + " |",
        md_table_separator(len(cols)),
    ]


def _section_file_meta(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    m = report.meta
    L.append(f"## {t('metadata.heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("metadata.table_columns", lang))
    L.append(f"| {t('metadata.row_filename', lang)} | `{m.file_name}` |")
    L.append(
        f"| {t('metadata.row_filesize', lang)} | "
        + t("metadata.row_filesize_value", lang,
            bytes_fmt=f"{m.file_size_bytes:,}", kb=f"{m.file_size_bytes/1024:.1f}")
        + " |"
    )
    L.append(f"| {t('metadata.row_sha256', lang)} | `{m.sha256}` |")
    if report.sanitized:
        L.append(
            f"| {t('metadata.row_sanitize', lang)} | "
            f"{t('metadata.sanitize_value', lang)} |"
        )
    L.append("")
    return L


def _section_basic_stats(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    b = report.basic_stats
    L.append(f"## {t('basic_stats.heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("basic_stats.table_columns", lang))
    L.append(f"| {t('basic_stats.row_sheet_count', lang)} | {b.sheet_count} |")
    L.append(
        f"| {t('basic_stats.row_sheet_visibility', lang)} | "
        f"{b.sheet_count_visible} / {b.sheet_count_hidden} / {b.sheet_count_very_hidden} |"
    )
    L.append(f"| {t('basic_stats.row_cells_nonempty', lang)} | {b.cell_count_nonempty:,} |")
    L.append(f"| {t('basic_stats.row_cells_formula', lang)} | {b.cell_count_formula:,} |")
    L.append(f"| {t('basic_stats.row_unique_values', lang)} | {b.cell_count_unique_values:,} |")
    L.append(f"| {t('basic_stats.row_named_ranges', lang)} | {b.named_range_count} |")
    L.append(f"| {t('basic_stats.row_cf', lang)} | {b.conditional_formatting_count} |")
    L.append(f"| {t('basic_stats.row_dv', lang)} | {b.data_validation_count} |")
    L.append(f"| {t('basic_stats.row_vba_modules', lang)} | {b.vba_module_count} |")
    L.append(f"| {t('basic_stats.row_vba_lines', lang)} | {b.vba_total_lines:,} |")
    L.append(f"| {t('basic_stats.row_parse_errors', lang)} | {b.parse_errors_count} |")
    L.append("")
    return L


def _section_sheets(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('sheets.heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("sheets.table_columns", lang))
    dash = t("common.dash", lang)
    for s in report.sheets:
        L.append(
            f"| `{_md_escape(s.name)}` | {s.state} | {s.rows_used} | {s.cols_used} | "
            f"{s.cells_nonempty} | {s.cells_formula} | {s.max_ref or dash} | "
            f"{s.conditional_formatting_count} | {s.data_validation_count} |"
        )
    L.append("")
    return L


def _section_named_ranges(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('named_ranges.heading', lang)}")
    L.append("")
    if not report.named_ranges:
        L.append(f"_{t('named_ranges.none', lang)}_")
    else:
        L.extend(_md_table_header_lines("named_ranges.table_columns", lang))
        for nr in report.named_ranges:
            L.append(f"| `{_md_escape(nr.name)}` | `{_md_escape(nr.scope)}` | `{_md_escape(nr.ref)}` |")
    L.append("")
    return L


def _section_complexity(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    c = report.complexity
    L.append(f"## {t('complexity.heading', lang)}")
    L.append("")
    L.append(f"**{t('complexity.total_label', lang, total=c.total)}**")
    L.append("")
    L.extend(_md_table_header_lines("complexity.table_columns", lang))
    sub = c.sub_scores
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        bar = _sub_score_bar(val)
        L.append(f"| {key.replace('_',' ')} | {val}/20 | `{bar}` | {_md_escape(c.rationale[key])} |")
    L.append("")
    return L


def _format_pillar_value_cell(p, lang: str = DEFAULT_LANG) -> str:
    """Inline value display for the Pillar table Value column (D1)."""
    v = (getattr(p, "value", "") or "").strip()
    kind = getattr(p, "value_kind", "")
    if not v:
        return f"_{t('pillar.value_empty', lang)}_"
    if kind == "formula":
        return f"`{_md_escape(v)}`{t('pillar.value_formula_suffix', lang)}"
    return f"`{_md_escape(v)}`"


def _format_pillar_label_cell(p, lang: str = DEFAULT_LANG) -> str:
    """Combine row_header/col_header/named_range into one column."""
    parts: list = []
    rh = (getattr(p, "row_header", "") or "").strip()
    ch = (getattr(p, "col_header", "") or "").strip()
    nr = (getattr(p, "named_range", "") or "").strip()
    if nr:
        parts.append(f"{t('pillar.label_named_range_prefix', lang)} `{_md_escape(nr)}`")
    if rh:
        parts.append(f"{t('pillar.label_row_prefix', lang)} `{_md_escape(rh[:25])}`")
    if ch and ch != rh:
        parts.append(f"{t('pillar.label_col_prefix', lang)} `{_md_escape(ch[:25])}`")
    return "; ".join(parts) if parts else t("common.dash", lang)


def _section_pillars(report, top_n: int = None, with_drilldown: bool = True,
                     anchor_id: str = "", lang: str = DEFAULT_LANG) -> list:
    """Pillar table — D1 column update: Cell | Value | Label | Members | Fan-in | …

    `top_n`: when set, only render top-N rows (used by Top Impact Findings).
              When None, render all (used by Reference Appendix 8.1).
    """
    L: list = []
    L.append(f"## {t('pillar.heading', lang)}")
    L.append("")
    L.append(f"_{t('pillar.intro', lang)}_")
    L.append("")
    L.append(f"_{t('pillar.what_this_means', lang)}_")
    L.append("")
    if not report.pillars:
        threshold = report.methodology['logic_depth_thresholds']['pillar-fanin-min']
        L.append(f"_{t('pillar.none_at_threshold', lang, threshold=threshold)}_")
        L.append("")
        return L

    rows = report.pillars if top_n is None else report.pillars[:top_n]
    L.extend(_md_table_header_lines("pillar.table_columns", lang))
    for i, p in enumerate(rows, start=1):
        sheets_str = ", ".join(f"`{s}`" for s in p.affected_sheets[:5])
        if len(p.affected_sheets) > 5:
            sheets_str += t("common.more_suffix", lang, n=len(p.affected_sheets) - 5)
        narrative = _render_pillar_narrative(p, lang)
        L.append(
            f"| {i} | `{_md_escape(p.location)}` | {_format_pillar_value_cell(p, lang)} | "
            f"{_format_pillar_label_cell(p, lang)} | {p.member_count} | {p.fan_in} | "
            f"{sheets_str} | {p.pillar_kind} | {_md_escape(narrative)} |"
        )
    if top_n is not None and len(report.pillars) > top_n:
        L.append("")
        L.append(f"_{t('pillar.more_pillars', lang, n=len(report.pillars) - top_n)}_")
    L.append("")
    if with_drilldown:
        L.append(f"**{t('pillar.drilldown_heading', lang)}**")
        L.append("")
        for p in rows[:5]:
            if p.member_count > 1:
                prefix = t("pillar.drilldown_group_label", lang,
                           location=p.location, fan_in=p.fan_in,
                           member_count=p.member_count)
            else:
                prefix = t("pillar.drilldown_single_label", lang,
                           location=p.location, fan_in=p.fan_in)
            L.append(f"- **{prefix}**:")
            for d in p.sample_dependents:
                L.append(f"    - `{d}`")
            if p.member_count > 1 and p.member_refs:
                preview = ", ".join(f"`{r}`" for r in p.member_refs[:3])
                extra = (t("common.more_suffix", lang, n=len(p.member_refs) - 3)
                         if len(p.member_refs) > 3 else "")
                L.append(f"    - {t('pillar.group_members_label', lang, preview=preview, extra=extra)}")
        L.append("")
    return L


def _section_anomalies(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('anomaly.heading', lang)}")
    L.append("")
    L.append(f"_{t('anomaly.intro', lang)}_")
    L.append("")
    if not report.anomalies:
        L.append(f"_{t('anomaly.none_detected', lang)}_")
        L.append("")
        return L
    L.extend(_md_table_header_lines("anomaly.table_columns", lang))
    for i, a in enumerate(report.anomalies, start=1):
        locs = ", ".join(f"`{loc}`" for loc in a.outlier_locations[:5])
        if len(a.outlier_locations) > 5:
            locs += t("common.more_suffix", lang, n=len(a.outlier_locations) - 5)
        narrative = _render_anomaly_narrative(a, lang)
        L.append(
            f"| {i} | `{_md_escape(a.cluster_pattern_sample)}` | "
            f"{a.cluster_size} | `{_md_escape(a.mode_value)}` "
            f"({a.mode_count}/{a.cluster_size}) | "
            f"`{_md_escape(a.outlier_value)}` ({a.outlier_count}/{a.cluster_size}) | "
            f"{locs} | {a.confidence} | {_md_escape(narrative)} |"
        )
    L.append("")
    return L


def _section_smells(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('smells.heading', lang)}")
    L.append("")
    types_count = len({s.smell_type for s in report.smells})
    L.append(f"_{t('smells.summary', lang, total=len(report.smells), types=types_count)}_")
    L.append("")
    by_type: dict = defaultdict(list)
    for s in report.smells:
        by_type[s.smell_type].append(s)
    for st in SMELL_TYPES:
        items = by_type.get(st, [])
        L.append(f"### {t('smells.subheading', lang, type=st, count=len(items))}")
        L.append("")
        if not items:
            L.append(f"_{t('smells.no_findings_threshold', lang)}_")
            L.append("")
            continue
        L.extend(_md_table_header_lines("smells.table_columns", lang))
        for s in items[:20]:
            L.append(
                f"| `{_md_escape(s.location)}` | {s.metric:g} | {s.severity} | "
                f"{s.confidence} | {_md_escape(s.evidence)} |"
            )
        if len(items) > 20:
            L.append("")
            L.append(f"_{t('smells.more_findings', lang, n=len(items) - 20)}_")
        L.append("")
    return L


def _section_magic_index(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('magic_index.heading', lang)}")
    L.append("")
    if not report.magic_numbers:
        L.append(f"_{t('magic_index.none', lang)}_")
    else:
        L.extend(_md_table_header_lines("magic_index.table_columns", lang))
        for mn in report.magic_numbers:
            L.append(
                f"| `{_md_escape(mn.value)}` | {mn.occurrence_count} | "
                f"`{_md_escape(mn.first_location)}` | {mn.location_kind} | "
                f"`{_md_escape(mn.sample_context)}` |"
            )
    L.append("")
    return L


def _section_vba(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('vba.heading', lang)}")
    L.append("")
    if not report.vba_modules:
        L.append(f"_{t('vba.no_modules_simple', lang)}_")
        L.append("")
        return L
    cls_by_name = {c.module_name: c for c in (report.vba_classifications or [])}
    dash = t("common.dash", lang)
    yes = t("common.yes", lang)
    no = t("common.no", lang)
    L.extend(_md_table_header_lines("vba.modules_table_columns", lang))
    for vm in report.vba_modules:
        n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
        n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
        cls = cls_by_name.get(vm.name)
        inferred = cls.inferred_type if cls else dash
        cls_conf = cls.confidence if cls else dash
        reads = ", ".join(f"`{s}`" for s in (cls.reads_sheets if cls else [])) or dash
        writes = ", ".join(f"`{s}`" for s in (cls.writes_sheets if cls else [])) or dash
        ext_call = (yes if (cls and cls.external_calls) else no) if cls else dash
        L.append(
            f"| `{_md_escape(vm.name)}` | {vm.type} | {vm.line_count} | "
            f"{n_sub} | {n_func} | **{inferred}** | {cls_conf} | "
            f"{_md_escape(reads)} | {_md_escape(writes)} | {ext_call} | "
            f"{yes if vm.has_on_error_resume_next else no} |"
        )
    L.append("")
    L.append(f"**{t('vba.details_heading', lang)}**")
    L.append("")
    L.extend(_md_table_header_lines("vba.details_table_columns", lang))
    for vm in report.vba_modules:
        ext = ", ".join(vm.external_keywords) or dash
        ranges = len(vm.range_literals)
        cls = cls_by_name.get(vm.name)
        rationale = cls.rationale if cls else dash
        L.append(
            f"| `{_md_escape(vm.name)}` | {_md_escape(ext)} | {ranges} | "
            f"{_md_escape(rationale)} |"
        )
    L.append("")
    return L


def _section_risks(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    r = report.risk_indicators
    dash = t("common.dash", lang)
    L.append(f"## {t('risks.heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("risks.table_columns", lang))
    L.append(
        f"| {t('risks.row_hidden', lang)} | {len(r.hidden_sheets)} "
        f"(`{', '.join(r.hidden_sheets) if r.hidden_sheets else dash}`) |"
    )
    L.append(
        f"| {t('risks.row_very_hidden', lang)} | {len(r.very_hidden_sheets)} "
        f"(`{', '.join(r.very_hidden_sheets) if r.very_hidden_sheets else dash}`) |"
    )
    L.append(f"| {t('risks.row_cross_sheet', lang)} | {r.cross_sheet_reference_count} |")
    L.append(f"| {t('risks.row_errors', lang)} | {len(r.cells_with_errors)} |")
    L.append(f"| {t('risks.row_external', lang)} | {len(r.external_workbook_references)} |")
    L.append(f"| {t('risks.row_circular', lang)} | {len(r.circular_reference_suspects)} |")
    L.append(f"| {t('risks.row_parse', lang)} | {len(r.parse_errors)} |")
    L.append("")
    if r.cells_with_errors:
        L.append(f"**{t('risks.errors_heading', lang)}**")
        L.append("")
        L.extend(_md_table_header_lines("risks.errors_table_columns", lang))
        for ec in r.cells_with_errors[:20]:
            L.append(f"| `{_md_escape(ec['sheet'])}` | `{ec['ref']}` | `{ec['error_token']}` |")
        L.append("")
    if r.external_workbook_references:
        L.append(f"**{t('risks.external_heading', lang)}**")
        L.append("")
        for ext in r.external_workbook_references[:20]:
            L.append(f"- `{_md_escape(ext)}`")
        L.append("")
    if r.circular_reference_suspects:
        L.append(f"**{t('risks.circular_heading', lang)}**")
        L.append("")
        for cs in r.circular_reference_suspects[:20]:
            L.append(f"- `{_md_escape(cs)}`")
        L.append("")
    return L


def _section_diagrams(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('diagrams.sheet_dataflow_heading', lang)}")
    L.append("")
    L.append(f"_{t('diagrams.sheet_dataflow_intro', lang)}_")
    L.append("")
    L.append("```mermaid")
    L.append(_build_sheet_dataflow_mermaid(report))
    L.append("```")
    L.append("")

    L.append(f"## {t('vba.classification_overview', lang)}")
    L.append("")
    L.append(f"_{t('vba.diagram_overview_intro', lang)}_")
    L.append("")
    L.append("```mermaid")
    L.append(_build_vba_classification_mermaid(report))
    L.append("```")
    L.append("")

    L.append(f"## {t('diagrams.pillar_impact_heading', lang)}")
    L.append("")
    L.append(f"_{t('diagrams.pillar_impact_intro', lang, top_n=_DIAG3_MAX_PILLARS)}_")
    L.append("")
    L.append("```mermaid")
    L.append(_build_pillar_impact_mermaid(report))
    L.append("```")
    L.append("")
    return L


def _section_methodology(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('methodology.heading', lang)}")
    L.append("")
    libs = report.methodology["library_versions"]
    L.append(f"- {t('methodology.engine', lang, openpyxl=libs['openpyxl'], oletools=libs['oletools'], formulas=libs['formulas'])}")
    thr = report.methodology["smell_thresholds"]
    L.append(f"- {t('methodology.smell_thresholds', lang, mr=thr['multiple-references'], lcc=thr['long-calculation-chain'], cc=thr['conditional-complexity'], mo=thr['multiple-operations'], df=thr['duplicated-formulas'])}")
    ld = report.methodology["logic_depth_thresholds"]
    L.append(f"- {t('methodology.logic_depth', lang, pfm=ld['pillar-fanin-min'], ptn=ld['pillar-top-n'], acms=ld['anomaly-cluster-min-size'], aof=ld['anomaly-outlier-fraction'])}")
    L.append(f"- {t('methodology.pillar_dedupe', lang, top_n=ld['pillar-top-n'])}")
    cats = ", ".join(f"`{c}`" for c in report.methodology["vba_classifier_categories"])
    L.append(f"- {t('methodology.vba_categories', lang, categories=cats)}")
    L.append(f"- {t('methodology.confidence_semantics', lang)}")
    L.append(f"- {t('methodology.domain_detector', lang)}")
    nums = ", ".join(f"`{n}`" for n in sorted(TRIVIAL_NUMBERS))
    L.append(f"- {t('methodology.trivial_numbers', lang, numbers=nums)}")
    L.append(f"- {t('methodology.reliability', lang)}")
    if report.sanitized:
        L.append(f"- {t('methodology.sanitize_active', lang)}")
    L.append("")
    L.append("---")
    L.append("")
    L.append(f"_{t('methodology.end_of_report', lang)}_")
    return L


# ---------------------------------------------------------------------------
# Round-3 new section builders — Workflow / Data Flow / VBA Walkthrough /
# Domain Findings / Top Impact Findings / Glossary.
# ---------------------------------------------------------------------------

def _section_workflow_guide(report, lang: str = DEFAULT_LANG) -> list:
    """D3: Workflow Guide — operational walkthrough.

    Derived from xl/drawings + VBA structure. When NO buttons + NO event
    handlers, the section reports "no buttons detected" gracefully.
    """
    L: list = []
    L.append(f"## {t('workflow.heading', lang)}")
    L.append("")
    L.append(f"_{t('workflow.intro', lang)}_")
    L.append("")
    L.append(f"_{t('workflow.what_this_means', lang)}_")
    L.append("")
    wf = (getattr(report, "workflow", None) or {})
    steps = wf.get("steps") or []

    if not steps:
        L.append(f"_{t('workflow.no_buttons', lang)}_")
        L.append("")
        return L

    for s in steps:
        L.append(f"<!-- LLM-AUGMENT: workflow-step:{s.order} -->")
        L.append(t("workflow.step_line", lang,
                   order=s.order,
                   label=_md_escape(s.label or s.sub_name),
                   sheet=_md_escape(s.sheet),
                   module_name=_md_escape(s.module_name),
                   sub_name=_md_escape(s.sub_name)))
        if s.reads_sheets:
            joined = ", ".join(f"`{_md_escape(r)}`" for r in s.reads_sheets[:5])
            extra = (t("common.more_suffix", lang, n=len(s.reads_sheets) - 5)
                     if len(s.reads_sheets) > 5 else "")
            L.append(f"- {t('workflow.reads_label', lang)}: {joined}{extra}")
        if s.writes_sheets:
            joined = ", ".join(f"`{_md_escape(w)}`" for w in s.writes_sheets[:5])
            extra = (t("common.more_suffix", lang, n=len(s.writes_sheets) - 5)
                     if len(s.writes_sheets) > 5 else "")
            L.append(f"- {t('workflow.writes_label', lang)}: {joined}{extra}")
        if s.calls:
            joined = ", ".join(f"`{_md_escape(c)}`" for c in s.calls[:5])
            extra = (t("common.more_suffix", lang, n=len(s.calls) - 5)
                     if len(s.calls) > 5 else "")
            L.append(f"- {t('workflow.calls_label', lang)}: {joined}{extra}")
        L.append("")

    # Mermaid sequence diagram
    L.append(f"**{t('workflow.sequence_heading', lang)}**:")
    L.append("")
    L.append("```mermaid")
    L.append(_build_workflow_mermaid(report))
    L.append("```")
    L.append("")
    return L


def _section_data_flow_story(report, lang: str = DEFAULT_LANG) -> list:
    """D5: Data Flow Story — per-sheet prose paragraph (top 8 by density).

    Renders BEFORE the existing Mermaid sheet-flow diagram (which becomes
    "evidence" for this prose).
    """
    L: list = []
    L.append(f"## {t('data_flow.heading', lang)}")
    L.append("")
    L.append(f"_{t('data_flow.intro', lang)}_")
    L.append("")
    L.append(f"_{t('data_flow.what_this_means', lang)}_")
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
        L.append("### " + t("data_flow.sheet_header", lang,
                            name=_md_escape(s.name),
                            state=s.state,
                            rows=s.rows_used,
                            cols=s.cols_used,
                            cells_nonempty=s.cells_nonempty))
        # Role inference
        n_in = sum(c for _, c in incoming.get(s.name, []))
        n_out = sum(c for _, c in outgoing.get(s.name, []))
        n_form = s.cells_formula
        n_vba_writes = len(vba_writes_by_sheet.get(s.name, []))

        if n_form == 0 and n_in == 0 and n_vba_writes == 0:
            role = t("data_flow.role_input", lang)
        elif n_form == 0 and (n_in > 0 or n_vba_writes > 0):
            role = t("data_flow.role_derived_no_formulas", lang)
        elif n_form > 0 and n_out > n_in:
            role = t("data_flow.role_aggregator", lang)
        elif n_form > 0:
            role = t("data_flow.role_computed", lang)
        else:
            role = t("data_flow.role_mixed", lang)
        L.append(t("data_flow.role_line", lang, role=role))

        # Sources / consumers
        if outgoing.get(s.name):
            srcs = ", ".join(f"`{src}` ({cnt})" for src, cnt in outgoing[s.name][:4])
            L.append(t("data_flow.sources_line", lang, sources=srcs))
        if incoming.get(s.name):
            tgts = ", ".join(f"`{tgt}` ({cnt})" for tgt, cnt in incoming[s.name][:4])
            L.append(t("data_flow.consumers_line", lang, consumers=tgts))
        if n_vba_writes > 0:
            mods = vba_writes_by_sheet[s.name][:3]
            mods_str = ", ".join(f"`{m}`" for m in mods)
            extra = (t("common.more_suffix", lang, n=n_vba_writes - 3)
                     if n_vba_writes > 3 else "")
            L.append(t("data_flow.vba_writes_line", lang, modules=mods_str, extra=extra))
        if pillar_count_by_sheet.get(s.name):
            L.append(t("data_flow.pillar_count_line", lang,
                       count=pillar_count_by_sheet[s.name]))
        if err_cells_by_sheet.get(s.name):
            ec = err_cells_by_sheet[s.name][:1][0]
            L.append(t("data_flow.manual_override_line", lang,
                       ref=ec['ref'], error_token=ec['error_token']))
        L.append("")

    if len(report.sheets) > len(top_sheets):
        L.append(f"_{t('data_flow.more_sheets', lang, n=len(report.sheets) - len(top_sheets))}_")
        L.append("")

    # The schematic diagram
    L.append(f"**{t('data_flow.diagram_heading', lang)}**:")
    L.append("")
    L.append(f"_{t('data_flow.diagram_intro', lang)}_")
    L.append("")
    L.append("```mermaid")
    L.append(_build_sheet_dataflow_mermaid(report))
    L.append("```")
    L.append("")
    return L


def _section_top_impact_findings(report, lang: str = DEFAULT_LANG) -> list:
    """Top Impact Findings — Top-5 each of pillars / anomalies / smells / risks.

    SHORT and READABLE — appendix has the full data.
    """
    L: list = []
    L.append(f"## {t('top_impact.heading', lang)}")
    L.append("")
    L.append(f"_{t('top_impact.intro', lang)}_")
    L.append("")

    # Pillars (top 5)
    L.append(f"### {t('top_impact.pillars_heading', lang)}")
    L.append("")
    if not report.pillars:
        L.append(f"_{t('top_impact.no_pillars', lang)}_")
    else:
        L.extend(_md_table_header_lines("pillar.table_columns_short", lang))
        for i, p in enumerate(report.pillars[:5], start=1):
            sheets_str = ", ".join(f"`{s}`" for s in p.affected_sheets[:3])
            if len(p.affected_sheets) > 3:
                sheets_str += f" (+{len(p.affected_sheets)-3})"
            L.append(
                f"| {i} | `{_md_escape(p.location)}` | {_format_pillar_value_cell(p, lang)} | "
                f"{_format_pillar_label_cell(p, lang)} | {p.fan_in} | {sheets_str} |"
            )
        L.append("")
        L.append(f"_{t('top_impact.pillars_see_full', lang)}_")
    L.append("")

    # Anomalies (top 5)
    L.append(f"### {t('top_impact.anomalies_heading', lang)}")
    L.append("")
    if not report.anomalies:
        L.append(f"_{t('top_impact.no_anomalies', lang)}_")
    else:
        L.extend(_md_table_header_lines("anomaly.table_columns_short", lang))
        for i, a in enumerate(report.anomalies[:5], start=1):
            locs = ", ".join(f"`{loc}`" for loc in a.outlier_locations[:3])
            if len(a.outlier_locations) > 3:
                locs += t("common.more_suffix", lang, n=len(a.outlier_locations) - 3)
            L.append(
                f"| {i} | {a.cluster_size} | `{_md_escape(a.mode_value)}` "
                f"({a.mode_count}/{a.cluster_size}) | "
                f"`{_md_escape(a.outlier_value)}` ({a.outlier_count}/{a.cluster_size}) | "
                f"{locs} | {a.confidence} |"
            )
    L.append("")

    # Top smells (top 5 by metric)
    L.append(f"### {t('top_impact.smells_heading', lang)}")
    L.append("")
    L.append(f"_{t('top_impact.smells_what_this_means', lang)}_")
    L.append("")
    if not report.smells:
        L.append(f"_{t('top_impact.no_smells', lang)}_")
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
        L.extend(_md_table_header_lines("smells.short_table_columns", lang))
        for i, s in enumerate(top_smells, start=1):
            L.append(
                f"| {i} | `{s.smell_type}` | `{_md_escape(s.location)}` | "
                f"{s.metric:g} | {s.severity} | {_md_escape(s.evidence)} |"
            )
        L.append("")
        L.append(f"_{t('top_impact.smells_full', lang)}_")
    L.append("")

    # Top risks
    L.append(f"### {t('top_impact.risks_heading', lang)}")
    L.append("")
    r = report.risk_indicators
    risk_lines: list = []
    if r.very_hidden_sheets:
        risk_lines.append(t("top_impact.risk_very_hidden", lang,
                            n=len(r.very_hidden_sheets),
                            names=', '.join(r.very_hidden_sheets)))
    if r.hidden_sheets:
        risk_lines.append(t("top_impact.risk_hidden", lang,
                            n=len(r.hidden_sheets),
                            names=', '.join(r.hidden_sheets)))
    if r.cells_with_errors:
        risk_lines.append(t("top_impact.risk_errors", lang,
                            n=len(r.cells_with_errors)))
    if r.external_workbook_references:
        risk_lines.append(t("top_impact.risk_external", lang,
                            n=len(r.external_workbook_references)))
    if r.circular_reference_suspects:
        risk_lines.append(t("top_impact.risk_circular", lang,
                            n=len(r.circular_reference_suspects)))
    if not risk_lines:
        L.append(f"_{t('top_impact.no_risks', lang)}_")
    else:
        for line in risk_lines:
            L.append(f"- {line}")
    L.append("")
    return L


def _section_vba_walkthrough(report, lang: str = DEFAULT_LANG) -> list:
    """D4: VBA Module Walkthrough — prose narration in call order.

    Each module gets a paragraph; modules with no callers flagged as
    "possibly dead code". LLM-AUGMENT markers reserve the slot for Track B.
    """
    L: list = []
    L.append(f"## {t('vba.heading_walkthrough', lang)}")
    L.append("")
    L.append(f"_{t('vba.walkthrough_intro', lang)}_")
    L.append("")
    L.append(f"_{t('vba.walkthrough_what_this_means', lang)}_")
    L.append("")
    narratives = report.vba_narratives or []
    if not narratives:
        L.append(f"_{t('vba.no_modules', lang)}_")
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
        L.append(f"_{t('vba.no_entry_caveat', lang)}_")
        L.append("")
        L.append(f"_{t('vba.many_dead_caveat', lang)}_")
        L.append("")
        for narr in narratives_to_show:
            L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
            L.append("### " + t("vba.module_heading", lang,
                                name=_md_escape(narr.module_name),
                                inferred_type=narr.inferred_type,
                                line_count=narr.line_count))
            L.append("")
            L.append(_render_vba_narrative(narr, lang))
            L.append("")
        if truly_dead:
            L.append("### " + t("vba.dead_code_heading", lang, n=len(truly_dead)))
            L.append("")
            sample = truly_dead[:8]
            for narr in sample:
                L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
                L.append("- " + t("vba.dead_module_bullet", lang,
                                  name=_md_escape(narr.module_name),
                                  inferred_type=narr.inferred_type,
                                  line_count=narr.line_count,
                                  sub_func_count=narr.sub_count + narr.func_count))
            if len(truly_dead) > len(sample):
                L.append("- " + f"_{t('vba.more_dead', lang, n=len(truly_dead) - len(sample))}_")
            L.append("")
        return L

    reachable_show = reachable[:REACHABLE_MAX]
    for narr in reachable_show:
        L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
        L.append("### " + t("vba.module_heading", lang,
                            name=_md_escape(narr.module_name),
                            inferred_type=narr.inferred_type,
                            line_count=narr.line_count))
        L.append("")
        L.append(_render_vba_narrative(narr, lang))
        L.append("")

    if len(reachable) > REACHABLE_MAX:
        L.append(f"_{t('vba.more_reachable', lang, n=len(reachable) - REACHABLE_MAX)}_")
        L.append("")

    if unreachable:
        L.append("### " + t("vba.dead_code_heading_alt", lang, n=len(unreachable)))
        L.append("")
        L.append(f"_{t('vba.dead_code_intro', lang)}_")
        L.append("")
        sample = unreachable[:8]
        for narr in sample:
            L.append(f"<!-- LLM-AUGMENT: vba-narration:{narr.module_name} -->")
            L.append("- " + t("vba.dead_module_bullet", lang,
                              name=_md_escape(narr.module_name),
                              inferred_type=narr.inferred_type,
                              line_count=narr.line_count,
                              sub_func_count=narr.sub_count + narr.func_count))
        if len(unreachable) > len(sample):
            L.append("- " + f"_{t('vba.more_dead', lang, n=len(unreachable) - len(sample))}_")
        L.append("")
    return L


def _section_domain_findings(report, lang: str = DEFAULT_LANG) -> list:
    """D7: Domain-Specific Findings — only present when ≥1 template fired."""
    L: list = []
    matches = (report.domain_template_matches or [])
    # Filter to high/medium confidence; low-confidence is noise here
    matches = [m for m in matches if m.confidence in ("high", "medium")]
    if not matches:
        return []  # caller skips section

    L.append(f"## {t('domain_findings.heading', lang)}")
    L.append("")
    L.append(f"_{t('domain_findings.intro', lang)}_")
    L.append("")
    L.append(f"_{t('domain_findings.what_this_means', lang)}_")
    L.append("")

    for m in matches:
        L.append(f"<!-- LLM-AUGMENT: domain-method:{m.template_key} -->")
        L.append("### " + t("domain_findings.template_heading", lang,
                            name=m.business_friendly_name,
                            confidence=m.confidence))
        L.append("")
        if m.matched_keywords:
            kw_str = ", ".join(f"`{k}`" for k in m.matched_keywords)
            L.append(t("domain_findings.matched_keywords", lang, keywords=kw_str))
            L.append("")

        if m.sheet_role_hits:
            L.append(t("domain_findings.expected_roles_found", lang))
            for role, sheets in m.sheet_role_hits:
                L.append("- " + t("domain_findings.role_present", lang,
                                  role=role, sheets=sheets))
            L.append("")
        if m.sheet_role_misses:
            L.append(t("domain_findings.expected_roles_missing", lang))
            for role in m.sheet_role_misses[:5]:
                L.append("- " + t("domain_findings.role_missing", lang, role=role))
            L.append("")

        if m.hardcode_risk_hits:
            L.append(t("domain_findings.hardcode_risks_heading", lang))
            for label, evidence in m.hardcode_risk_hits:
                L.append("- " + t("domain_findings.hardcode_risk_item", lang,
                                  label=label, evidence=evidence))
            L.append("")
        else:
            L.append(t("domain_findings.hardcode_risks_none", lang))
            L.append("")

        if m.method_hits:
            L.append(t("domain_findings.method_hits_heading", lang))
            for label, evidence in m.method_hits:
                L.append("- " + t("domain_findings.method_hits_item", lang,
                                  label=label, evidence=evidence))
            L.append("")
    return L


_GLOSSARY_KEYS = (
    "anomaly", "audit", "byoa", "column_block", "complexity_score",
    "confidence", "vba_classes", "dead_suspected", "domain",
    "duplicated_formulas", "fan_in", "formula_relay", "hermans_smells",
    "incoming_outgoing", "llm_augment", "magic_number", "on_error",
    "pillar", "sanitize", "smell", "tier", "track_a_b", "very_hidden",
    "workflow_step",
)


def _glossary_term_def(slug: str, lang: str) -> tuple:
    """Return (term, definition) for a glossary slug, parsing the catalog
    `term||definition` value format."""
    raw = t(f"glossary.{slug}", lang)
    if "||" in raw:
        term, defn = raw.split("||", 1)
        return term.strip(), defn.strip()
    return raw, ""


def _section_glossary(report, lang: str = DEFAULT_LANG) -> list:
    """D8: Glossary — alphabetical plain-language definitions."""
    L: list = []
    L.append(f"## {t('glossary.heading', lang)}")
    L.append("")
    L.append(f"_{t('glossary.intro', lang)}_")
    L.append("")
    items: list = [_glossary_term_def(slug, lang) for slug in _GLOSSARY_KEYS]
    items.sort(key=lambda kv: kv[0].lower())
    for term, defn in items:
        L.append(f"- **{_md_escape(term)}** — {_md_escape(defn)}")
    L.append("")
    return L


def _section_executive_summary_round3(report, lang: str = DEFAULT_LANG) -> list:
    """Round-3 executive summary — 3-5 manager-readable headlines.

    Replaces the prior more verbose exec summary with a tightly-filtered
    pyramid top.
    """
    L: list = []
    c = report.complexity
    domain = getattr(report, "domain_hint", None)

    L.append(f"## {t('exec_summary.heading', lang)}")
    L.append("")
    # Headline: complexity + plain-language rendition
    if c.total >= 80:
        complexity_tier = t("exec_summary.complexity_tier_top", lang)
    elif c.total >= 50:
        complexity_tier = t("exec_summary.complexity_tier_moderate", lang)
    elif c.total >= 20:
        complexity_tier = t("exec_summary.complexity_tier_manageable", lang)
    else:
        complexity_tier = t("exec_summary.complexity_tier_small", lang)
    L.append("- " + t("exec_summary.complexity_line", lang,
                      total=c.total, tier=complexity_tier))

    # Top pillar
    if report.pillars:
        p = report.pillars[0]
        if p.member_count > 1:
            L.append("- " + t("exec_summary.most_referenced_group", lang,
                              location=p.location, member_count=p.member_count,
                              fan_in=p.fan_in))
        else:
            value_phrase = ""
            if p.value:
                value_phrase = t("exec_summary.value_phrase", lang, value=p.value[:30])
            L.append("- " + t("exec_summary.single_most_impactful", lang,
                              location=p.location, value_phrase=value_phrase,
                              fan_in=p.fan_in,
                              affected_sheet_count=p.affected_sheet_count))

    # Top smell or anomaly
    if report.anomalies:
        a = report.anomalies[0]
        L.append("- " + t("exec_summary.top_anomaly", lang,
                          cluster_size=a.cluster_size,
                          outlier_value=a.outlier_value,
                          mode_value=a.mode_value,
                          outlier_location=a.outlier_locations[0],
                          confidence=a.confidence))
    elif report.smells:
        top_smell = max(report.smells, key=lambda s: s.metric)
        L.append("- " + t("exec_summary.top_smell", lang,
                          smell_type=top_smell.smell_type,
                          location=top_smell.location,
                          metric=f"{top_smell.metric:g}",
                          severity=top_smell.severity))

    # Domain
    if domain is not None and domain.domain != "unknown":
        kw_str = ", ".join(domain.matched_keywords[:5])
        L.append("- " + t("exec_summary.detected_domain", lang,
                          domain=domain.domain,
                          confidence=domain.confidence,
                          keywords=kw_str))

    # Workflow
    wf = (getattr(report, "workflow", None) or {})
    n_buttons = len(wf.get("buttons", []) or [])
    n_events = len(wf.get("event_handlers", []) or [])
    if n_buttons or n_events:
        L.append("- " + t("exec_summary.entry_points", lang,
                          n_buttons=n_buttons, n_events=n_events))
    else:
        L.append("- " + t("exec_summary.no_entry_points", lang))

    # Risks
    r = report.risk_indicators
    risks_short = []
    if r.very_hidden_sheets:
        risks_short.append(t("exec_summary.risk_very_hidden_short", lang,
                             n=len(r.very_hidden_sheets)))
    if r.hidden_sheets:
        risks_short.append(t("exec_summary.risk_hidden_short", lang,
                             n=len(r.hidden_sheets)))
    if r.cells_with_errors:
        risks_short.append(t("exec_summary.risk_errors_short", lang,
                             n=len(r.cells_with_errors)))
    if r.external_workbook_references:
        risks_short.append(t("exec_summary.risk_external_short", lang,
                             n=len(r.external_workbook_references)))
    if risks_short:
        L.append("- " + t("exec_summary.risk_flags", lang,
                          risks="; ".join(risks_short)))

    # Sub-score table (compact)
    L.append("")
    L.append(f"**{t('exec_summary.subscores_heading', lang)}**")
    L.append("")
    L.extend(_md_table_header_lines("complexity.short_columns", lang))
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

def _section_appendix_intro(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"## {t('appendix.heading', lang)}")
    L.append("")
    L.append(f"_{t('appendix.intro', lang)}_")
    L.append("")
    return L


# VBA classification summary table (for D6 — replaces the unreadable diagram
# with a class-level table at the top, then a bounded mini-diagram).
def _section_vba_classification_summary(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"### {t('vba.classification_overview', lang)}")
    L.append("")
    classifications = report.vba_classifications or []
    if not classifications:
        L.append(f"_{t('vba.no_modules', lang)}_")
        L.append("")
        return L
    by_class: dict = defaultdict(list)
    for c in classifications:
        by_class[c.inferred_type].append(c)
    # Build module -> LOC lookup
    loc_by_name = {vm.name: vm.line_count for vm in report.vba_modules}

    L.extend(_md_table_header_lines("vba.classification_table_columns", lang))
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
            sample += t("common.more_suffix", lang, n=len(items) - 3)
        L.append(f"| `{cls}` | {len(items)} | {total_loc:,} | {sample} |")
    for cls, items in sorted(by_class.items()):
        if cls in seen:
            continue
        total_loc = sum(loc_by_name.get(c.module_name, 0) for c in items)
        sample = ", ".join(f"`{c.module_name}`" for c in items[:3])
        if len(items) > 3:
            sample += t("common.more_suffix", lang, n=len(items) - 3)
        L.append(f"| `{cls}` | {len(items)} | {total_loc:,} | {sample} |")
    L.append("")
    L.append(f"**{t('vba.classification_mini_diagram_heading', lang)}**")
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


def _exec_summary_lines(report, lang: str = DEFAULT_LANG) -> list:
    """Build the executive summary block. Returns markdown lines."""
    return _section_executive_summary_round3(report, lang)


def render_markdown(report, lang: str = DEFAULT_LANG) -> str:
    """Round-3 pyramid layout — selective narrative > exhaustive dump."""
    return _assemble_markdown(report, lang)


def _section_pillar_impact_diagram_only(report, lang: str = DEFAULT_LANG) -> list:
    """Standalone Pillar impact diagram (H3 under appendix)."""
    L: list = []
    L.append(f"### {t('diagrams.pillar_impact_heading', lang)}")
    L.append("")
    L.append(f"_{t('diagrams.pillar_impact_intro', lang, top_n=_DIAG3_MAX_PILLARS)}_")
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


def _section_appendix_pillars_full(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"### {t('appendix.pillars_heading', lang)}")
    L.append("")
    if not report.pillars:
        L.append(f"_{t('appendix.pillars_none_at_threshold', lang)}_")
        L.append("")
        return L
    L.append(f"_{t('appendix.pillars_intro', lang)}_")
    L.append("")
    L.extend(_md_table_header_lines("pillar.table_columns", lang))
    for i, p in enumerate(report.pillars, start=1):
        sheets_str = ", ".join(f"`{s}`" for s in p.affected_sheets[:5])
        if len(p.affected_sheets) > 5:
            sheets_str += t("common.more_suffix", lang, n=len(p.affected_sheets) - 5)
        narrative = _render_pillar_narrative(p, lang)
        L.append(
            f"| {i} | `{_md_escape(p.location)}` | {_format_pillar_value_cell(p, lang)} | "
            f"{_format_pillar_label_cell(p, lang)} | {p.member_count} | {p.fan_in} | "
            f"{sheets_str} | {p.pillar_kind} | {_md_escape(narrative)} |"
        )
    L.append("")
    return L


def _section_appendix_smells_full(report, lang: str = DEFAULT_LANG) -> list:
    L: list = []
    L.append(f"### {t('appendix.smells_heading', lang)}")
    L.append("")
    types_count = len({s.smell_type for s in report.smells})
    L.append(f"_{t('appendix.smells_summary', lang, n=len(report.smells), types=types_count)}_")
    L.append("")
    by_type: dict = defaultdict(list)
    for s in report.smells:
        by_type[s.smell_type].append(s)
    for st in SMELL_TYPES:
        items = by_type.get(st, [])
        L.append(f"#### {t('appendix.smells_subheading', lang, type=st, count=len(items))}")
        L.append("")
        if not items:
            L.append(f"_{t('smells.no_findings_threshold', lang)}_")
            L.append("")
            continue
        L.extend(_md_table_header_lines("smells.table_columns", lang))
        for s in items[:20]:
            L.append(
                f"| `{_md_escape(s.location)}` | {s.metric:g} | {s.severity} | "
                f"{s.confidence} | {_md_escape(s.evidence)} |"
            )
        if len(items) > 20:
            L.append("")
            L.append(f"_{t('smells.more_findings', lang, n=len(items) - 20)}_")
        L.append("")
    return L


def _assemble_markdown(report, lang: str = DEFAULT_LANG) -> str:
    """Round-3 pyramid assembly.

    1. Cover (filename, audit version, sanitize banner)
    2. Executive Summary (manager-readable)
    3. Workflow Guide
    4. Data Flow Story
    5. Top Impact Findings (Top-N each)
    6. VBA Module Walkthrough
    7. Domain-Specific Findings (when domain detected)
    8. Reference Appendix (H3 sub-sections — full data)
    9. Glossary
    10. Methodology
    """
    L: list = []
    m = report.meta
    from . import __version__ as _pkg_version
    # 1. Cover
    headline_complexity = report.complexity.total
    L.append(
        f"# {t('cover.title', lang)} — `{m.file_name}` "
        f"({t('cover.audit_version', lang, version=_pkg_version)})"
    )
    L.append("")

    if report.sanitized:
        L.append(f"> {t('cover.sanitized_banner', lang)}")
        L.append(">")
    # cover.headline includes "complexity {complexity}, ..." — we pass the
    # bold-wrapped "**N/100**" so the markdown shows it bold.
    headline_md = t(
        "cover.headline", lang,
        complexity=f"**{headline_complexity}/100**",
        pillar_count=f"**{len(report.pillars)}**",
        smell_count=f"**{len(report.smells)}**",
    )
    L.append(f"> {headline_md}")
    L.append(">")
    L.append(f"> {t('cover.subline_1', lang)}")
    L.append(f"> {t('cover.subline_2', lang)}")
    L.append("")

    # Build TOC dynamically (skip Domain-Specific Findings if not detected)
    toc_headers = list(_TOP_LEVEL_HEADERS_FOR_TOC)
    if not _domain_findings_present(report):
        toc_headers = [h for h in toc_headers if h != "Domain-Specific Findings"]

    # 2. Executive Summary
    L.extend(_exec_summary_lines(report, lang))

    # TOC (right after exec). The header IS translated, but we store the
    # English anchors in the TOC links so the cross-link to per-section
    # H2 (whose anchor we compute from the EN slug) still resolves. To keep
    # the markdown link working, we use the EN-anchored slug — but display
    # the translated header text.
    L.append(f"## {t('exec_summary.toc_heading', lang)}")
    L.append("")
    for header_en in toc_headers:
        anchor = _slugify_anchor(header_en)
        # Translate the section heading text via lookup map
        header_display = _toc_header_display(header_en, lang)
        L.append(f"- [{header_display}](#{anchor})")
    L.append("")

    # 3. Workflow Guide
    L.extend(_section_workflow_guide(report, lang))

    # 4. Data Flow Story
    L.extend(_section_data_flow_story(report, lang))

    # 5. Top Impact Findings
    L.extend(_section_top_impact_findings(report, lang))

    # 6. VBA Module Walkthrough
    L.extend(_section_vba_walkthrough(report, lang))

    # 7. Domain-Specific Findings (only when present)
    if _domain_findings_present(report):
        L.extend(_section_domain_findings(report, lang))

    # 8. Reference Appendix (full data tables, H3-level)
    L.extend(_section_appendix_intro(report, lang))

    # 8.1 Pillar table (full)
    L.extend(_section_appendix_pillars_full(report, lang))

    # 8.2 Smell catalog (full)
    L.extend(_section_appendix_smells_full(report, lang))

    # 8.3 Sheets table
    L.append(f"### {t('appendix.sheets_heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("sheets.table_columns", lang))
    dash = t("common.dash", lang)
    for s in report.sheets:
        L.append(
            f"| `{_md_escape(s.name)}` | {s.state} | {s.rows_used} | {s.cols_used} | "
            f"{s.cells_nonempty} | {s.cells_formula} | {s.max_ref or dash} | "
            f"{s.conditional_formatting_count} | {s.data_validation_count} |"
        )
    L.append("")

    # 8.4 Named ranges
    L.append(f"### {t('appendix.named_ranges_heading', lang)}")
    L.append("")
    if not report.named_ranges:
        L.append(f"_{t('named_ranges.none', lang)}_")
    else:
        L.extend(_md_table_header_lines("named_ranges.table_columns", lang))
        for nr in report.named_ranges:
            L.append(f"| `{_md_escape(nr.name)}` | `{_md_escape(nr.scope)}` | `{_md_escape(nr.ref)}` |")
    L.append("")

    # 8.5 Magic-number index
    L.append(f"### {t('appendix.magic_heading', lang)}")
    L.append("")
    if not report.magic_numbers:
        L.append(f"_{t('magic_index.none', lang)}_")
    else:
        L.extend(_md_table_header_lines("magic_index.table_columns", lang))
        for mn in report.magic_numbers:
            L.append(
                f"| `{_md_escape(mn.value)}` | {mn.occurrence_count} | "
                f"`{_md_escape(mn.first_location)}` | {mn.location_kind} | "
                f"`{_md_escape(mn.sample_context)}` |"
            )
    L.append("")

    # 8.6 Risk indicators (full)
    r = report.risk_indicators
    L.append(f"### {t('appendix.risks_heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("risks.table_columns", lang))
    L.append(
        f"| {t('risks.row_hidden', lang)} | {len(r.hidden_sheets)} "
        f"(`{', '.join(r.hidden_sheets) if r.hidden_sheets else dash}`) |"
    )
    L.append(
        f"| {t('risks.row_very_hidden', lang)} | {len(r.very_hidden_sheets)} "
        f"(`{', '.join(r.very_hidden_sheets) if r.very_hidden_sheets else dash}`) |"
    )
    L.append(f"| {t('risks.row_cross_sheet', lang)} | {r.cross_sheet_reference_count} |")
    L.append(f"| {t('risks.row_errors', lang)} | {len(r.cells_with_errors)} |")
    L.append(f"| {t('risks.row_external', lang)} | {len(r.external_workbook_references)} |")
    L.append(f"| {t('risks.row_circular', lang)} | {len(r.circular_reference_suspects)} |")
    L.append(f"| {t('risks.row_parse', lang)} | {len(r.parse_errors)} |")
    L.append("")
    if r.cells_with_errors:
        L.append(f"**{t('risks.errors_heading', lang)}**")
        L.append("")
        L.extend(_md_table_header_lines("risks.errors_table_columns", lang))
        for ec in r.cells_with_errors[:20]:
            L.append(f"| `{_md_escape(ec['sheet'])}` | `{ec['ref']}` | `{ec['error_token']}` |")
        L.append("")
    if r.external_workbook_references:
        L.append(f"**{t('risks.external_heading', lang)}**")
        L.append("")
        for ext in r.external_workbook_references[:20]:
            L.append(f"- `{_md_escape(ext)}`")
        L.append("")
    if r.circular_reference_suspects:
        L.append(f"**{t('risks.circular_heading', lang)}**")
        L.append("")
        for cs in r.circular_reference_suspects[:20]:
            L.append(f"- `{_md_escape(cs)}`")
        L.append("")

    # 8.7 Complexity score breakdown
    c = report.complexity
    L.append(f"### {t('appendix.complexity_heading', lang)}")
    L.append("")
    L.append(f"**{t('complexity.total_label', lang, total=c.total)}**")
    L.append("")
    L.extend(_md_table_header_lines("complexity.table_columns", lang))
    sub = c.sub_scores
    for key in sorted(c.rationale.keys()):
        val = getattr(sub, key)
        bar = _sub_score_bar(val)
        L.append(f"| {key.replace('_',' ')} | {val}/20 | `{bar}` | {_md_escape(c.rationale[key])} |")
    L.append("")

    # 8.8 VBA classification + full table
    L.append(f"### {t('appendix.vba_heading', lang)}")
    L.append("")
    classifications = report.vba_classifications or []
    if not classifications:
        L.append(f"_{t('vba.no_modules', lang)}_")
        L.append("")
    else:
        by_class: dict = defaultdict(list)
        for cls_obj in classifications:
            by_class[cls_obj.inferred_type].append(cls_obj)
        loc_by_name = {vm.name: vm.line_count for vm in report.vba_modules}
        L.extend(_md_table_header_lines("vba.classification_table_columns", lang))
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
                sample += t("common.more_suffix", lang, n=len(items) - 3)
            L.append(f"| `{cls}` | {len(items)} | {total_loc:,} | {sample} |")
        for cls, items in sorted(by_class.items()):
            if cls in seen:
                continue
            total_loc = sum(loc_by_name.get(c.module_name, 0) for c in items)
            sample = ", ".join(f"`{c.module_name}`" for c in items[:3])
            if len(items) > 3:
                sample += t("common.more_suffix", lang, n=len(items) - 3)
            L.append(f"| `{cls}` | {len(items)} | {total_loc:,} | {sample} |")
        L.append("")
        L.append(f"**{t('vba.classification_mini_diagram_heading', lang)}**")
        L.append("")
        L.append("```mermaid")
        L.append(_build_vba_classification_mermaid(report))
        L.append("```")
        L.append("")
    # Full VBA modules table
    L.append(f"#### {t('vba.full_modules_table_heading', lang)}")
    L.append("")
    yes = t("common.yes", lang)
    no = t("common.no", lang)
    if not report.vba_modules:
        L.append(f"_{t('vba.no_modules_simple', lang)}_")
    else:
        cls_by_name = {c.module_name: c for c in (report.vba_classifications or [])}
        L.extend(_md_table_header_lines("vba.modules_table_columns", lang))
        for vm in report.vba_modules:
            n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
            n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
            cls = cls_by_name.get(vm.name)
            inferred = cls.inferred_type if cls else dash
            cls_conf = cls.confidence if cls else dash
            reads = ", ".join(f"`{s}`" for s in (cls.reads_sheets if cls else [])) or dash
            writes = ", ".join(f"`{s}`" for s in (cls.writes_sheets if cls else [])) or dash
            ext_call = (yes if (cls and cls.external_calls) else no) if cls else dash
            L.append(
                f"| `{_md_escape(vm.name)}` | {vm.type} | {vm.line_count} | "
                f"{n_sub} | {n_func} | **{inferred}** | {cls_conf} | "
                f"{_md_escape(reads)} | {_md_escape(writes)} | {ext_call} | "
                f"{yes if vm.has_on_error_resume_next else no} |"
            )
    L.append("")

    # 8.9 Pillar impact diagram (kept as evidence)
    L.append(f"### {t('appendix.diagram_heading', lang)}")
    L.append("")
    L.append(f"_{t('diagrams.pillar_impact_intro', lang, top_n=_DIAG3_MAX_PILLARS)}_")
    L.append("")
    L.append("```mermaid")
    L.append(_build_pillar_impact_mermaid(report))
    L.append("```")
    L.append("")

    # 8.10 File metadata + basic stats
    L.append(f"### {t('appendix.metadata_heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("metadata.table_columns", lang))
    L.append(f"| {t('metadata.row_filename', lang)} | `{m.file_name}` |")
    L.append(
        f"| {t('metadata.row_filesize', lang)} | "
        + t("metadata.row_filesize_value", lang,
            bytes_fmt=f"{m.file_size_bytes:,}", kb=f"{m.file_size_bytes/1024:.1f}")
        + " |"
    )
    L.append(f"| {t('metadata.row_sha256', lang)} | `{m.sha256}` |")
    if report.sanitized:
        L.append(
            f"| {t('metadata.row_sanitize', lang)} | "
            f"{t('metadata.sanitize_value', lang)} |"
        )
    L.append("")

    # 8.11 Basic statistics (the original table)
    b = report.basic_stats
    L.append(f"### {t('appendix.basic_stats_heading', lang)}")
    L.append("")
    L.extend(_md_table_header_lines("basic_stats.table_columns", lang))
    L.append(f"| {t('basic_stats.row_sheet_count', lang)} | {b.sheet_count} |")
    L.append(
        f"| {t('basic_stats.row_sheet_visibility', lang)} | "
        f"{b.sheet_count_visible} / {b.sheet_count_hidden} / {b.sheet_count_very_hidden} |"
    )
    L.append(f"| {t('basic_stats.row_cells_nonempty', lang)} | {b.cell_count_nonempty:,} |")
    L.append(f"| {t('basic_stats.row_cells_formula', lang)} | {b.cell_count_formula:,} |")
    L.append(f"| {t('basic_stats.row_unique_values', lang)} | {b.cell_count_unique_values:,} |")
    L.append(f"| {t('basic_stats.row_named_ranges', lang)} | {b.named_range_count} |")
    L.append(f"| {t('basic_stats.row_cf', lang)} | {b.conditional_formatting_count} |")
    L.append(f"| {t('basic_stats.row_dv', lang)} | {b.data_validation_count} |")
    L.append(f"| {t('basic_stats.row_vba_modules', lang)} | {b.vba_module_count} |")
    L.append(f"| {t('basic_stats.row_vba_lines', lang)} | {b.vba_total_lines:,} |")
    L.append(f"| {t('basic_stats.row_parse_errors', lang)} | {b.parse_errors_count} |")
    L.append("")

    # 9. Glossary
    L.extend(_section_glossary(report, lang))

    # 10. Methodology
    L.extend(_section_methodology(report, lang))

    return "\n".join(L) + "\n"


# Map of EN top-level header -> i18n heading key (for TOC display + HTML toc).
_TOC_HEADER_KEY_MAP = {
    "Executive Summary": "exec_summary.heading",
    "Workflow Guide": "workflow.heading",
    "Data Flow Story": "data_flow.heading",
    "Top Impact Findings": "top_impact.heading",
    "VBA Module Walkthrough": "vba.heading_walkthrough",
    "Domain-Specific Findings": "domain_findings.heading",
    "Reference Appendix": "appendix.heading",
    "Glossary": "glossary.heading",
    "Methodology": "methodology.heading",
}


def _toc_header_display(header_en: str, lang: str) -> str:
    """Translate a TOC header from English to the target language using the
    appropriate i18n heading key. Falls back to the English text."""
    key = _TOC_HEADER_KEY_MAP.get(header_en)
    if key is None:
        return header_en
    return t(key, lang)


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


def _html_section_open(title: str, anchor_override: str = "") -> str:
    """Build a `<section><h2>` opening tag.

    `anchor_override`: when provided, use this instead of slugifying the
    title. Used so a translated title (Chinese characters) still gets a
    stable EN-derived anchor for the HTML id.
    """
    anchor = anchor_override or _slugify_anchor(title)
    if not anchor:
        # Fallback when title is non-ASCII and no override: derive a stable
        # slug via simple hash so the section is still uniquely addressable.
        anchor = "section-" + str(abs(hash(title)) % 10**8)
    return f'<section class="audit-section" id="{anchor}"><h2>{_h(title)}</h2>'


def _html_section_close() -> str:
    return "</section>"


def _html_exec_summary(report, lang: str = DEFAULT_LANG) -> str:
    c = report.complexity
    domain = getattr(report, "domain_hint", None)
    findings = _headline_findings(report, lang)

    parts: list = []
    parts.append('<section class="audit-section exec-summary" id="executive-summary">')
    parts.append(f'<h2>{_h(t("exec_summary.heading", lang))}</h2>')
    parts.append(
        f"<p><strong>{_h(t('exec_summary.complexity_label', lang))}</strong>: "
        f"{c.total} / 100</p>"
    )

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
    cols = split_pipe_columns(t("complexity.short_columns", lang))
    parts.append(_html_table(cols, rows))

    parts.append(f"<p><strong>{_h(t('exec_summary.top_findings_label', lang))}</strong></p>")
    if findings:
        cards: list = []
        for f in findings:
            cls = "callout-card"
            if (f.startswith("**Pillar")
                    or f.startswith("**Systemrelevante")
                    or f.startswith("**支柱")):
                cls = "callout-card callout-pillar"
            elif (f.startswith("**Anomaly")
                    or f.startswith("**Anomalie")
                    or f.startswith("**异常")):
                cls = "callout-card callout-anomaly"
            cards.append(f'<div class="{cls}">{_inline_md_to_html(f)}</div>')
        parts.append("".join(cards))
    else:
        parts.append(f"<p><em>{_h(t('exec_summary.no_findings', lang))}</em></p>")

    if domain is not None:
        if domain.domain == "unknown":
            parts.append("<p>" + _inline_md_to_html(
                t("exec_summary.unknown_domain", lang)) + "</p>")
        else:
            kw_str = ", ".join(domain.matched_keywords)
            parts.append("<p>" + _inline_md_to_html(
                t("exec_summary.detected_domain_html", lang,
                  domain=domain.domain,
                  confidence=domain.confidence,
                  keywords=kw_str)) + "</p>")
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


def _html_methodology(report, lang: str = DEFAULT_LANG) -> str:
    out = [_html_section_open(t("methodology.heading", lang),
                              anchor_override="methodology")]
    libs = report.methodology["library_versions"]
    out.append("<ul><li>" + _inline_md_to_html(t(
        "methodology.engine", lang,
        openpyxl=libs['openpyxl'], oletools=libs['oletools'],
        formulas=libs['formulas'])) + "</li>")
    thr = report.methodology["smell_thresholds"]
    out.append("<li>" + _inline_md_to_html(t(
        "methodology.smell_thresholds", lang,
        mr=thr['multiple-references'],
        lcc=thr['long-calculation-chain'],
        cc=thr['conditional-complexity'],
        mo=thr['multiple-operations'],
        df=thr['duplicated-formulas'])).replace(
        "≥", "&ge;") + "</li>")
    ld = report.methodology["logic_depth_thresholds"]
    out.append("<li>" + _inline_md_to_html(t(
        "methodology.logic_depth", lang,
        pfm=ld['pillar-fanin-min'], ptn=ld['pillar-top-n'],
        acms=ld['anomaly-cluster-min-size'],
        aof=ld['anomaly-outlier-fraction'])).replace(
        "≥", "&ge;").replace("≤", "&le;") + "</li>")
    out.append("<li>" + _inline_md_to_html(t(
        "methodology.pillar_dedupe_short", lang)) + "</li>")
    cats = ", ".join(f"`{c}`" for c in report.methodology["vba_classifier_categories"])
    out.append("<li>" + _inline_md_to_html(t(
        "methodology.vba_categories", lang, categories=cats)) + "</li>")
    out.append("<li>" + _inline_md_to_html(
        t("methodology.confidence_semantics", lang)) + "</li>")
    out.append("<li>" + _inline_md_to_html(
        t("methodology.domain_detector_short", lang)) + "</li>")
    nums = ", ".join(f"`{n}`" for n in sorted(TRIVIAL_NUMBERS))
    out.append("<li>" + _inline_md_to_html(
        t("methodology.trivial_numbers", lang, numbers=nums)) + "</li>")
    out.append("<li>" + _inline_md_to_html(
        t("methodology.reliability", lang)).replace("→", "&rarr;") + "</li>")
    if report.sanitized:
        out.append("<li>" + _inline_md_to_html(
            t("methodology.sanitize_active", lang)) + "</li>")
    out.append("</ul>")
    out.append(f"<hr><p><em>{_h(t('methodology.end_of_report', lang))}</em></p>")
    out.append(_html_section_close())
    return "".join(out)


def _html_toc(report, lang: str = DEFAULT_LANG) -> str:
    """Round-3 TOC: only the top-level pyramid headers, in order."""
    headers = [h for h in _TOP_LEVEL_HEADERS_FOR_TOC]
    if not _domain_findings_present(report):
        headers = [h for h in headers if h != "Domain-Specific Findings"]
    out: list = [f'<nav class="toc audit-section" id="toc"><h2>{_h(t("exec_summary.toc_heading", lang))}</h2><ul>']
    for header in headers:
        anchor = _slugify_anchor(header)
        display = _toc_header_display(header, lang)
        out.append(f'<li><a href="#{anchor}">{_h(display)}</a></li>')
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
    # Pre-build a reverse lookup so a translated H2 title can resolve back to
    # the EN section anchor (so HTML <section id="..."> stays stable across
    # languages and TOC links keep working).
    _heading_to_en_slug = {}
    for en_title, key in _TOC_HEADER_KEY_MAP.items():
        en_slug = _slugify_anchor(en_title)
        for lang_code in SUPPORTED_LANGS:
            translated = t(key, lang_code)
            _heading_to_en_slug[translated] = en_slug
            _heading_to_en_slug[en_title] = en_slug
    while i < n:
        line = lines[i]
        # H2
        if line.startswith("## "):
            title = line[3:].strip()
            # Prefer the EN-derived anchor for a known translated heading.
            anchor = _heading_to_en_slug.get(title) or _slugify_anchor(title)
            if not anchor:
                # Fallback: stable slug from any non-empty title hash
                anchor = "section-" + str(abs(hash(title)) % 10**8)
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


def _html_workflow_guide(report, lang: str = DEFAULT_LANG) -> str:
    return _md_lines_to_html(_section_workflow_guide(report, lang))


def _html_data_flow_story(report, lang: str = DEFAULT_LANG) -> str:
    return _md_lines_to_html(_section_data_flow_story(report, lang))


def _html_top_impact_findings(report, lang: str = DEFAULT_LANG) -> str:
    return _md_lines_to_html(_section_top_impact_findings(report, lang))


def _html_vba_walkthrough(report, lang: str = DEFAULT_LANG) -> str:
    return _md_lines_to_html(_section_vba_walkthrough(report, lang))


def _html_domain_findings(report, lang: str = DEFAULT_LANG) -> str:
    return _md_lines_to_html(_section_domain_findings(report, lang))


def _html_glossary(report, lang: str = DEFAULT_LANG) -> str:
    return _md_lines_to_html(_section_glossary(report, lang))


def _html_appendix(report, lang: str = DEFAULT_LANG) -> str:
    """Render the Reference Appendix using the existing tableized helpers
    (so the appendix retains rich-HTML tables — preferred over MD-derived)."""
    parts: list = []
    parts.append(f'<section class="audit-section" id="reference-appendix"><h2>{_h(t("appendix.heading", lang))}</h2>')
    parts.append(f"<p><em>{_h(t('appendix.intro', lang))}</em></p>")
    # 8.1 Pillar table (full) — reuse _html_pillars but rename heading
    parts.append(f"<h3>{_h(t('appendix.pillars_heading', lang))}</h3>")
    parts.append(_html_pillars_table_only(report, lang))
    # 8.2 Smells catalog
    parts.append(f"<h3>{_h(t('appendix.smells_heading', lang))}</h3>")
    parts.append(_html_smells_inner(report, lang))
    # 8.3 Sheets
    parts.append(f"<h3>{_h(t('appendix.sheets_heading', lang))}</h3>")
    parts.append(_html_sheets_inner(report, lang))
    # 8.4 Named ranges
    parts.append(f"<h3>{_h(t('appendix.named_ranges_heading', lang))}</h3>")
    parts.append(_html_named_ranges_inner(report, lang))
    # 8.5 Magic-number index
    parts.append(f"<h3>{_h(t('appendix.magic_heading', lang))}</h3>")
    parts.append(_html_magic_index_inner(report, lang))
    # 8.6 Risk indicators
    parts.append(f"<h3>{_h(t('appendix.risks_heading', lang))}</h3>")
    parts.append(_html_risks_inner(report, lang))
    # 8.7 Complexity breakdown
    parts.append(f"<h3>{_h(t('appendix.complexity_heading', lang))}</h3>")
    parts.append(_html_complexity_inner(report, lang))
    # 8.8 VBA classification + table
    parts.append(f"<h3>8.8 {_h(t('vba.classification_overview', lang))}</h3>")
    parts.append(_html_vba_summary_inner(report, lang))
    parts.append(f"<h4>{_h(t('vba.full_modules_table_heading', lang))}</h4>")
    parts.append(_html_vba_inner(report, lang))
    # 8.9 Pillar impact diagram
    parts.append(f"<h3>8.9 {_h(t('diagrams.pillar_impact_heading', lang))}</h3>")
    parts.append(f"<p><em>{_h(t('diagrams.pillar_impact_intro', lang, top_n=_DIAG3_MAX_PILLARS))}</em></p>")
    parts.append(f'<pre class="mermaid">{_h(_build_pillar_impact_mermaid(report))}</pre>')
    # 8.10 File metadata
    parts.append(f"<h3>{_h(t('appendix.metadata_heading', lang))}</h3>")
    parts.append(_html_file_metadata_inner(report, lang))
    # 8.11 Basic statistics
    parts.append(f"<h3>{_h(t('appendix.basic_stats_heading', lang))}</h3>")
    parts.append(_html_basic_stats_inner(report, lang))
    parts.append("</section>")
    return "".join(parts)


# Inner table helpers — strip the section wrapper so the appendix can place
# them under H3 sub-headers rather than separate H2 sections.
def _html_pillars_table_only(report, lang: str = DEFAULT_LANG) -> str:
    if not report.pillars:
        thr = report.methodology["logic_depth_thresholds"]["pillar-fanin-min"]
        # The original used "fan-in &ge; N" — translate via the markdown key
        # then escape to entity form.
        msg = t("pillar.none_at_threshold", lang, threshold=thr).replace("≥", "&ge;")
        return f"<p><em>{msg}</em></p>"
    rows = []
    for i, p in enumerate(report.pillars, start=1):
        sheets_str = ", ".join(_html_code(s) for s in p.affected_sheets[:5])
        if len(p.affected_sheets) > 5:
            sheets_str += t("common.more_suffix", lang, n=len(p.affected_sheets) - 5)
        # Value cell HTML
        v = (getattr(p, "value", "") or "").strip()
        val_html = (f"<em>{_h(t('pillar.value_empty', lang))}</em>"
                    if not v else f"<code>{_h(v[:30])}</code>")
        # Label cell HTML — same labels as the markdown row, prefixed via i18n
        label_parts = []
        if getattr(p, "named_range", ""):
            label_parts.append(
                f"{_h(t('pillar.label_named_range_prefix', lang))} "
                f"<code>{_h(p.named_range)}</code>"
            )
        if getattr(p, "row_header", ""):
            label_parts.append(
                f"{_h(t('pillar.label_row_prefix', lang))} "
                f"<code>{_h(p.row_header[:25])}</code>"
            )
        if getattr(p, "col_header", "") and p.col_header != getattr(p, "row_header", ""):
            label_parts.append(
                f"{_h(t('pillar.label_col_prefix', lang))} "
                f"<code>{_h(p.col_header[:25])}</code>"
            )
        label_html = "; ".join(label_parts) if label_parts else _h(t("common.dash", lang))
        narrative = _render_pillar_narrative(p, lang)
        rows.append([
            _h(i), _html_code(p.location), val_html, label_html,
            _h(p.member_count), _h(p.fan_in),
            sheets_str, _h(p.pillar_kind), _h(narrative),
        ])
    cols = split_pipe_columns(t("pillar.table_columns", lang))
    return _html_table(cols, rows)


def _html_smells_inner(report, lang: str = DEFAULT_LANG) -> str:
    types_count = len({s.smell_type for s in report.smells})
    out = [f"<p><em>{_h(t('appendix.smells_summary', lang, n=len(report.smells), types=types_count))}</em></p>"]
    by_type: dict = defaultdict(list)
    for s in report.smells:
        by_type[s.smell_type].append(s)
    for st in SMELL_TYPES:
        items = by_type.get(st, [])
        out.append(f"<h4><code>{_h(st)}</code> &mdash; {_h(len(items))} finding(s)</h4>")
        if not items:
            out.append(f"<p><em>{_h(t('smells.no_findings_threshold', lang))}</em></p>")
            continue
        rows = []
        for s in items[:20]:
            rows.append([
                _html_code(s.location), _h(f"{s.metric:g}"),
                _h(s.severity), _h(s.confidence), _h(s.evidence),
            ])
        cols = split_pipe_columns(t("smells.table_columns", lang))
        out.append(_html_table(cols, rows))
        if len(items) > 20:
            # Template contains backtick markdown — render it inline
            out.append(f"<p><em>{_inline_md_to_html(t('smells.more_findings', lang, n=len(items) - 20))}</em></p>")
    return "".join(out)


def _html_sheets_inner(report, lang: str = DEFAULT_LANG) -> str:
    rows = []
    dash = t("common.dash", lang)
    for s in report.sheets:
        rows.append([
            _html_code(s.name), _h(s.state), _h(s.rows_used), _h(s.cols_used),
            _h(s.cells_nonempty), _h(s.cells_formula),
            _h(s.max_ref or dash),
            _h(s.conditional_formatting_count), _h(s.data_validation_count),
        ])
    cols = split_pipe_columns(t("sheets.table_columns", lang))
    return _html_table(cols, rows)


def _html_named_ranges_inner(report, lang: str = DEFAULT_LANG) -> str:
    if not report.named_ranges:
        return f"<p><em>{_h(t('named_ranges.none', lang))}</em></p>"
    rows = [
        [_html_code(nr.name), _html_code(nr.scope), _html_code(nr.ref)]
        for nr in report.named_ranges
    ]
    cols = split_pipe_columns(t("named_ranges.table_columns", lang))
    return _html_table(cols, rows)


def _html_magic_index_inner(report, lang: str = DEFAULT_LANG) -> str:
    if not report.magic_numbers:
        return f"<p><em>{_h(t('magic_index.none', lang))}</em></p>"
    rows = [
        [
            _html_code(mn.value), _h(mn.occurrence_count),
            _html_code(mn.first_location), _h(mn.location_kind),
            _html_code(mn.sample_context),
        ]
        for mn in report.magic_numbers
    ]
    cols = split_pipe_columns(t("magic_index.table_columns", lang))
    return _html_table(cols, rows)


def _html_risks_inner(report, lang: str = DEFAULT_LANG) -> str:
    out: list = []
    r = report.risk_indicators
    dash = t("common.dash", lang)
    rows = [
        [t("risks.row_hidden", lang),
         _h(f"{len(r.hidden_sheets)} ({', '.join(r.hidden_sheets) if r.hidden_sheets else dash})")],
        [t("risks.row_very_hidden", lang),
         _h(f"{len(r.very_hidden_sheets)} ({', '.join(r.very_hidden_sheets) if r.very_hidden_sheets else dash})")],
        [t("risks.row_cross_sheet", lang), _h(r.cross_sheet_reference_count)],
        [t("risks.row_errors", lang), _h(len(r.cells_with_errors))],
        [t("risks.row_external", lang), _h(len(r.external_workbook_references))],
        [t("risks.row_circular", lang), _h(len(r.circular_reference_suspects))],
        [t("risks.row_parse", lang), _h(len(r.parse_errors))],
    ]
    cols = split_pipe_columns(t("risks.table_columns", lang))
    out.append(_html_table(cols, rows))
    if r.cells_with_errors:
        # The original used "Formula error cells (first 20)" without colon for h4
        eh = t("risks.errors_heading", lang).rstrip(":")
        out.append(f"<h4>{_h(eh)}</h4>")
        rows = [
            [_html_code(ec["sheet"]), _html_code(ec["ref"]), _html_code(ec["error_token"])]
            for ec in r.cells_with_errors[:20]
        ]
        ecols = split_pipe_columns(t("risks.errors_table_columns", lang))
        out.append(_html_table(ecols, rows))
    if r.external_workbook_references:
        eh = t("risks.external_heading", lang).rstrip(":")
        out.append(f"<h4>{_h(eh)}</h4><ul>")
        for ext in r.external_workbook_references[:20]:
            out.append(f"<li>{_html_code(ext)}</li>")
        out.append("</ul>")
    if r.circular_reference_suspects:
        ch = t("risks.circular_heading", lang).rstrip(":")
        out.append(f"<h4>{_h(ch)}</h4><ul>")
        for cs in r.circular_reference_suspects[:20]:
            out.append(f"<li>{_html_code(cs)}</li>")
        out.append("</ul>")
    return "".join(out)


def _html_complexity_inner(report, lang: str = DEFAULT_LANG) -> str:
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
    cols = split_pipe_columns(t("complexity.table_columns", lang))
    return (f"<p><strong>{_h(t('complexity.total_word', lang))}</strong>: "
            f"{_h(c.total)} / 100</p>"
            + _html_table(cols, rows))


def _html_vba_summary_inner(report, lang: str = DEFAULT_LANG) -> str:
    """Render VBA classification summary + the bounded mermaid diagram."""
    classifications = report.vba_classifications or []
    if not classifications:
        return f"<p><em>{_h(t('vba.no_modules', lang))}</em></p>"
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
            sample += t("common.more_suffix", lang, n=len(items) - 3)
        rows.append([_html_code(cls), _h(len(items)), _h(f"{total_loc:,}"), sample])
    for cls, items in sorted(by_class.items()):
        if cls in seen:
            continue
        total_loc = sum(loc_by_name.get(c.module_name, 0) for c in items)
        sample = ", ".join(_html_code(c.module_name) for c in items[:3])
        if len(items) > 3:
            sample += t("common.more_suffix", lang, n=len(items) - 3)
        rows.append([_html_code(cls), _h(len(items)), _h(f"{total_loc:,}"), sample])
    cols = split_pipe_columns(t("vba.classification_table_columns", lang))
    out = [_html_table(cols, rows)]
    h_text = t("vba.classification_mini_diagram_heading", lang).rstrip(":").replace(
        "≤", "&le;")
    out.append(f"<h4>{h_text}</h4>")
    out.append(f'<pre class="mermaid">{_h(_build_vba_classification_mermaid(report))}</pre>')
    return "".join(out)


def _html_vba_inner(report, lang: str = DEFAULT_LANG) -> str:
    if not report.vba_modules:
        return f"<p><em>{_h(t('vba.no_modules_simple', lang))}</em></p>"
    cls_by_name = {c.module_name: c for c in (report.vba_classifications or [])}
    rows = []
    dash = t("common.dash", lang)
    yes = t("common.yes", lang)
    no = t("common.no", lang)
    for vm in report.vba_modules:
        n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
        n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
        cls = cls_by_name.get(vm.name)
        inferred = cls.inferred_type if cls else dash
        cls_conf = cls.confidence if cls else dash
        reads = ", ".join(_html_code(s) for s in (cls.reads_sheets if cls else [])) or dash
        writes = ", ".join(_html_code(s) for s in (cls.writes_sheets if cls else [])) or dash
        ext_call = (yes if (cls and cls.external_calls) else no) if cls else dash
        rows.append([
            _html_code(vm.name), _h(vm.type), _h(vm.line_count),
            _h(n_sub), _h(n_func),
            f"<strong>{_h(inferred)}</strong>",
            _h(cls_conf), reads, writes, _h(ext_call),
            _h(yes if vm.has_on_error_resume_next else no),
        ])
    cols = split_pipe_columns(t("vba.modules_table_columns", lang))
    return _html_table(cols, rows)


def _html_file_metadata_inner(report, lang: str = DEFAULT_LANG) -> str:
    m = report.meta
    rows = [
        [t("metadata.row_filename", lang), _html_code(m.file_name)],
        [t("metadata.row_filesize", lang),
         _h(t("metadata.row_filesize_value", lang,
              bytes_fmt=f"{m.file_size_bytes:,}",
              kb=f"{m.file_size_bytes/1024:.1f}"))],
        [t("metadata.row_sha256", lang), _html_code(m.sha256)],
    ]
    if report.sanitized:
        # Use the same rendered phrasing as MD; convert **active** to <strong>
        rows.append([t("metadata.row_sanitize", lang),
                     _inline_md_to_html(t("metadata.sanitize_value", lang))])
    cols = split_pipe_columns(t("metadata.table_columns", lang))
    return _html_table(cols, rows)


def _html_basic_stats_inner(report, lang: str = DEFAULT_LANG) -> str:
    b = report.basic_stats
    rows = [
        [t("basic_stats.row_sheet_count", lang), _h(b.sheet_count)],
        [t("basic_stats.row_sheet_visibility", lang),
         _h(f"{b.sheet_count_visible} / {b.sheet_count_hidden} / {b.sheet_count_very_hidden}")],
        [t("basic_stats.row_cells_nonempty", lang), _h(f"{b.cell_count_nonempty:,}")],
        [t("basic_stats.row_cells_formula", lang), _h(f"{b.cell_count_formula:,}")],
        [t("basic_stats.row_unique_values", lang), _h(f"{b.cell_count_unique_values:,}")],
        [t("basic_stats.row_named_ranges", lang), _h(b.named_range_count)],
        [t("basic_stats.row_cf", lang), _h(b.conditional_formatting_count)],
        [t("basic_stats.row_dv", lang), _h(b.data_validation_count)],
        [t("basic_stats.row_vba_modules", lang), _h(b.vba_module_count)],
        [t("basic_stats.row_vba_lines", lang), _h(f"{b.vba_total_lines:,}")],
        [t("basic_stats.row_parse_errors", lang), _h(b.parse_errors_count)],
    ]
    cols = split_pipe_columns(t("basic_stats.table_columns", lang))
    return _html_table(cols, rows)


_MERMAID_CDN = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"


def render_html(report, mermaid_inline: bool = False,
                mermaid_inline_source: str = "",
                lang: str = DEFAULT_LANG) -> str:
    """Render the audit as a styled HTML page (round-3 pyramid layout)."""
    from . import __version__ as _pkg_version
    m = report.meta

    parts: list = []
    parts.append("<!DOCTYPE html>")
    parts.append(f'<html lang="{_h(lang)}"><head>')
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

    parts.append(f"<h1>{_h(t('cover.title', lang))} &mdash; <code>{_h(m.file_name)}</code> "
                 f"({_h(t('cover.audit_version', lang, version=_pkg_version))})</h1>")

    if report.sanitized:
        # Reuse the markdown banner template; convert ` ` and `**` markers
        # to HTML.
        banner_md = t("cover.sanitized_banner", lang)
        # Strip leading "🔒 " emoji for HTML version (we use entity)
        banner_inner = banner_md
        if banner_inner.startswith("🔒 "):
            banner_inner = banner_inner[len("🔒 "):]
        banner_html = _inline_md_to_html(banner_inner).replace(
            "&lt;redacted&gt;", "&lt;redacted&gt;"  # already entity-escaped via _inline
        )
        parts.append(f'<div class="banner-sanitize">&#128274; {banner_html}</div>')

    headline_complexity = report.complexity.total
    headline = t("cover.headline", lang,
                 complexity=f"<strong>{headline_complexity}/100</strong>",
                 pillar_count=f"<strong>{len(report.pillars)}</strong>",
                 smell_count=f"<strong>{len(report.smells)}</strong>")
    sub1 = t("cover.subline_1", lang)
    sub2 = t("cover.subline_2", lang)
    parts.append(f"<blockquote>{headline}<br>{sub1} {sub2}</blockquote>")

    # Pyramid order:
    parts.append(_html_exec_summary(report, lang))
    parts.append(_html_toc(report, lang))
    parts.append(_html_workflow_guide(report, lang))
    parts.append(_html_data_flow_story(report, lang))
    parts.append(_html_top_impact_findings(report, lang))
    parts.append(_html_vba_walkthrough(report, lang))
    if _domain_findings_present(report):
        parts.append(_html_domain_findings(report, lang))
    parts.append(_html_appendix(report, lang))
    parts.append(_html_glossary(report, lang))
    parts.append(_html_methodology(report, lang))

    parts.append("</body></html>")
    return "".join(parts) + "\n"
