"""Track B (BYOA) extract phase — build dossier + mega-prompt for the user's LLM.

Architectural promise: this module never calls an LLM, never opens a network
socket, and never imports any LLM SDK. It only PREPARES context the user can
paste into their own Copilot Chat / Claude Desktop / ChatGPT / etc.

Inputs:
    - The fully-built AuditReport (from audit.build_audit)
    - The rendered audit.md text (so we can pull the exact LLM-AUGMENT marker
      IDs the renderer emitted — no duplication of render.py logic)

Outputs:
    - dossier.json — structured workbook context (deterministic, sort_keys)
    - prompt.md — opinionated mega-prompt instructing the LLM what to write
      and forcing strict-JSON output (so ingest.py can parse it)

Idempotency: same AuditReport + same audit.md -> byte-identical dossier.json
and prompt.md every time. We sort dict keys, deduplicate marker IDs, and use
stable iteration order over the report's already-deterministic structures.
"""

from __future__ import annotations

import json
import re
from dataclasses import asdict
from pathlib import Path

# Field names skipped when we serialize report fragments into the dossier
# (mirrors render._EXCLUDED_FIELDS — keep these in sync).
_DOSSIER_EXCLUDED_FIELDS = frozenset({"_sheet_edges"})

_RE_MARKER = re.compile(r"<!--\s*LLM-AUGMENT:\s*([^\s][^>]*?)\s*-->")
# Stricter: marker must be alone on its line (mirrors the renderer's emit
# pattern). This prevents picking up the marker example in the Glossary
# (where the comment appears inline in a sentence).
_RE_MARKER_LINE = re.compile(r"^\s*<!--\s*LLM-AUGMENT:\s*([^\s][^>]*?)\s*-->\s*$",
                              re.MULTILINE)

# How many sample formulas to keep per sheet in the dossier. Enough for the
# LLM to recognize the pattern (Holt-Winters / SUMIFS-tile / VLOOKUP-tile)
# without bloating the file.
_SAMPLE_FORMULAS_PER_SHEET = 6
_SAMPLE_FORMULA_MAX_LEN = 220

# Top-N caps that match render.py's section caps (keeps dossier coherent
# with what the markdown actually shows).
_TOP_PILLARS = 10
_TOP_SMELLS = 10
_TOP_MAGIC_NUMBERS = 20
_TOP_ANOMALIES = 10
_VBA_SOURCE_MAX_CHARS_PER_MODULE = 12000  # safety cap for one truly huge module


# =============================================================================
# Dossier construction
# =============================================================================

def _strip_excluded(d):
    """Recursively drop excluded keys from a dataclass-derived dict."""
    if isinstance(d, dict):
        return {k: _strip_excluded(v)
                for k, v in d.items()
                if k not in _DOSSIER_EXCLUDED_FIELDS}
    if isinstance(d, (list, tuple)):
        return [_strip_excluded(x) for x in d]
    return d


def _dataclass_to_dict(obj):
    """Convert a dataclass to a JSON-friendly dict (or pass-through)."""
    if obj is None:
        return None
    if hasattr(obj, "__dataclass_fields__"):
        return _strip_excluded(asdict(obj))
    return obj


def _truncate_str(s: str, max_len: int) -> str:
    if not s:
        return s
    if len(s) <= max_len:
        return s
    return s[: max_len - 1].rstrip() + "…"


def _sample_formulas_for_sheet(sheet_name: str, cell_rows) -> list:
    """Pick up to N representative formulas from a sheet.

    Strategy: dedupe by formula pattern (the formula text itself with cell
    refs preserved is fine — the LLM benefits from seeing exact patterns)
    and take the longest distinct formulas first; ties broken by cell ref
    for determinism.
    """
    seen: set = set()
    out: list = []
    # Filter to this sheet's formula cells, sorted by ref for determinism
    cells = sorted(
        (cr for cr in cell_rows if cr.sheet == sheet_name and cr.formula),
        key=lambda cr: (-len(cr.formula or ""), cr.row, cr.col, cr.ref),
    )
    for cr in cells:
        f = cr.formula
        if f in seen:
            continue
        seen.add(f)
        out.append({
            "ref": cr.ref,
            "formula": _truncate_str(f, _SAMPLE_FORMULA_MAX_LEN),
        })
        if len(out) >= _SAMPLE_FORMULAS_PER_SHEET:
            break
    return out


def _build_workbook_meta(report) -> dict:
    dh = report.domain_hint
    domain_detected = ""
    domain_confidence = ""
    if dh is not None:
        domain_detected = getattr(dh, "domain", "") or ""
        domain_confidence = getattr(dh, "confidence", "") or ""
    return {
        "filename": report.meta.file_name,
        "file_size_bytes": report.meta.file_size_bytes,
        "sha256": report.meta.sha256,
        "sheet_count": report.basic_stats.sheet_count,
        "sheet_count_visible": report.basic_stats.sheet_count_visible,
        "sheet_count_hidden": report.basic_stats.sheet_count_hidden,
        "sheet_count_very_hidden": report.basic_stats.sheet_count_very_hidden,
        "cell_count_nonempty": report.basic_stats.cell_count_nonempty,
        "cell_count_formula": report.basic_stats.cell_count_formula,
        "vba_module_count": report.basic_stats.vba_module_count,
        "vba_total_lines": report.basic_stats.vba_total_lines,
        "complexity_total": report.complexity.total,
        "complexity_subscores": _dataclass_to_dict(report.complexity.sub_scores),
        "domain_detected": domain_detected,
        "domain_confidence": domain_confidence,
        "sanitized": bool(report.sanitized),
    }


def _build_sheets_section(report, cell_rows) -> list:
    """All sheets (not just top-N) — short structural summary + formula samples.

    Sample formulas are crucial: the LLM uses them to recognize methods
    (Holt-Winters, SUMIFS-tile, MRP netting, etc.). We cap at
    _SAMPLE_FORMULAS_PER_SHEET to keep dossier compact.
    """
    out: list = []
    for s in report.sheets:
        out.append({
            "name": s.name,
            "state": s.state,
            "rows": s.rows_used,
            "cols": s.cols_used,
            "cells_nonempty": s.cells_nonempty,
            "cells_formula": s.cells_formula,
            "max_ref": s.max_ref,
            "sample_formulas": _sample_formulas_for_sheet(s.name, cell_rows),
        })
    return out


def _build_vba_modules_section(report) -> list:
    """All VBA modules with FULL source — the LLM needs source to write
    semantic narratives. Capped per-module by _VBA_SOURCE_MAX_CHARS_PER_MODULE
    only for the rare 12000+ char monster.

    `called_by` and `calls` are derived from VbaNarrative records (which
    already computed them) so we don't re-run the call-graph here.
    """
    # Build module -> classification + narrative lookups
    cls_by_mod = {c.module_name: c for c in (report.vba_classifications or [])}
    narr_by_mod = {n.module_name: n for n in (report.vba_narratives or [])}

    out: list = []
    for vm in (report.vba_modules or []):
        cls = cls_by_mod.get(vm.name)
        narr = narr_by_mod.get(vm.name)
        src = vm.source_text or ""
        truncated = False
        if len(src) > _VBA_SOURCE_MAX_CHARS_PER_MODULE:
            src = src[:_VBA_SOURCE_MAX_CHARS_PER_MODULE]
            truncated = True
        out.append({
            "name": vm.name,
            "type": vm.type,
            "lines": vm.line_count,
            "sub_function_names": [sf.name for sf in vm.sub_functions],
            "external_keywords": list(vm.external_keywords or []),
            "has_on_error_resume_next": bool(vm.has_on_error_resume_next),
            "inferred_type": cls.inferred_type if cls else "mixed",
            "classification_confidence": cls.confidence if cls else "low",
            "reads_sheets": list(cls.reads_sheets) if cls else [],
            "writes_sheets": list(cls.writes_sheets) if cls else [],
            "called_by": list(narr.callers) if narr else [],
            "calls": list(narr.callees) if narr else [],
            "reachable_from_entry": bool(narr.reachable_from_entry) if narr else False,
            "source": src,
            "source_truncated": truncated,
        })
    return out


def _build_workflow_section(report) -> dict:
    wf = (getattr(report, "workflow", None) or {})
    steps = wf.get("steps") or []
    out_steps: list = []
    for s in steps:
        out_steps.append({
            "order": s.order,
            "label": s.label,
            "macro": s.macro,
            "sub_name": s.sub_name,
            "module_name": s.module_name,
            "sheet": s.sheet,
            "reads_sheets": list(s.reads_sheets),
            "writes_sheets": list(s.writes_sheets),
            "calls": list(s.calls),
            "source_kind": s.source_kind,
        })
    return {
        "no_buttons_detected": bool(wf.get("no_buttons_detected")),
        "steps": out_steps,
    }


def _build_pillars_top(report) -> list:
    out: list = []
    for p in (report.pillars or [])[:_TOP_PILLARS]:
        out.append({
            "location": p.location,
            "value": p.value,
            "value_kind": p.value_kind,
            "row_header": p.row_header,
            "col_header": p.col_header,
            "named_range": p.named_range,
            "fan_in": p.fan_in,
            "affected_sheets": list(p.affected_sheets),
            "pillar_kind": p.pillar_kind,
            "is_formula_itself": bool(p.is_formula_itself),
            "member_count": p.member_count,
        })
    return out


def _build_smells_top(report) -> list:
    out: list = []
    for s in (report.smells or [])[:_TOP_SMELLS]:
        out.append({
            "smell_type": s.smell_type,
            "location": s.location,
            "metric": s.metric,
            "severity": s.severity,
            "confidence": s.confidence,
            "evidence": s.evidence,
        })
    return out


def _build_magic_numbers_top(report) -> list:
    out: list = []
    for m in (report.magic_numbers or [])[:_TOP_MAGIC_NUMBERS]:
        out.append({
            "value": m.value,
            "occurrence_count": m.occurrence_count,
            "first_location": m.first_location,
            "sample_context": _truncate_str(m.sample_context, 200),
            "location_kind": m.location_kind,
            "confidence": m.confidence,
        })
    return out


def _build_anomalies_top(report) -> list:
    out: list = []
    for a in (report.anomalies or [])[:_TOP_ANOMALIES]:
        out.append({
            "cluster_pattern_sample": _truncate_str(a.cluster_pattern_sample, 200),
            "cluster_size": a.cluster_size,
            "mode_value": a.mode_value,
            "mode_count": a.mode_count,
            "outlier_value": a.outlier_value,
            "outlier_count": a.outlier_count,
            "outlier_locations": list(a.outlier_locations),
            "confidence": a.confidence,
        })
    return out


def _build_domain_templates(report) -> list:
    out: list = []
    for m in (report.domain_template_matches or []):
        out.append({
            "template_key": m.template_key,
            "business_friendly_name": m.business_friendly_name,
            "matched_keywords": list(m.matched_keywords),
            "sheet_role_hits": [list(t) for t in m.sheet_role_hits],
            "sheet_role_misses": list(m.sheet_role_misses),
            "method_hits": [list(t) for t in m.method_hits],
            "confidence": m.confidence,
        })
    return out


# =============================================================================
# Marker extraction + per-marker context
# =============================================================================

def _extract_markers(audit_md_text: str) -> list:
    """Return ordered list of unique marker IDs found in the audit.md.

    The renderer is the source of truth — we don't duplicate the logic that
    decides which markers to emit. We just read what it actually wrote.
    De-duplication preserves first-seen order (some markers can theoretically
    appear twice in degenerate cases — we ask the LLM once).
    """
    seen: set = set()
    ordered: list = []
    for m in _RE_MARKER_LINE.finditer(audit_md_text):
        marker_id = m.group(1).strip()
        if not marker_id or marker_id in seen:
            continue
        seen.add(marker_id)
        ordered.append(marker_id)
    return ordered


def _context_for_marker(marker_id: str, report) -> dict:
    """Build (context, ask) for one marker.

    `context` is a 1-2 sentence summary of WHAT the marker is about (so the
    LLM doesn't have to re-derive it). `ask` is the specific instruction
    for what kind of narrative to write.

    Markers are typed by the prefix before the first colon:
        workflow-step:<N>
        data-flow:<sheet>
        vba-narration:<module>
        domain-method:<template_key>
    Unknown prefixes get a generic ask.
    """
    if ":" not in marker_id:
        return {
            "id": marker_id,
            "context": "(no structured context available)",
            "ask": ("Write 2-4 sentences explaining the business meaning of "
                    "this section."),
        }
    prefix, rest = marker_id.split(":", 1)

    if prefix == "workflow-step":
        return _ctx_workflow_step(marker_id, rest, report)
    if prefix == "data-flow":
        return _ctx_data_flow(marker_id, rest, report)
    if prefix == "vba-narration":
        return _ctx_vba_narration(marker_id, rest, report)
    if prefix == "domain-method":
        return _ctx_domain_method(marker_id, rest, report)

    return {
        "id": marker_id,
        "context": f"Unknown marker type: {prefix}",
        "ask": "Write 2-4 sentences of business-relevant narrative.",
    }


def _ctx_workflow_step(marker_id: str, rest: str, report) -> dict:
    try:
        order = int(rest)
    except ValueError:
        order = -1
    wf = (getattr(report, "workflow", None) or {})
    step = next((s for s in (wf.get("steps") or []) if s.order == order), None)
    if step is None:
        ctx = f"Workflow step #{order} (details not found in report)."
    else:
        reads = ", ".join(step.reads_sheets[:5]) or "(none)"
        writes = ", ".join(step.writes_sheets[:5]) or "(none)"
        calls = ", ".join(step.calls[:5]) or "(none)"
        ctx = (
            f"Step {step.order}: user clicks button '{step.label}' on sheet "
            f"'{step.sheet}', bound to {step.module_name}.{step.sub_name}. "
            f"Reads sheets: {reads}. Writes sheets: {writes}. Calls subs: {calls}. "
            f"Source kind: {step.source_kind}."
        )
    return {
        "id": marker_id,
        "context": ctx,
        "ask": (
            "In 2-3 sentences, explain what this step ACHIEVES for the "
            "business — not what it does mechanically. If you can identify "
            "the operation (e.g. 'recomputes MRP net requirements', 'rebuilds "
            "shift schedule', 'pulls forecast from web service'), name it."
        ),
    }


def _ctx_data_flow(marker_id: str, rest: str, report) -> dict:
    sheet = next((s for s in report.sheets if s.name == rest), None)
    if sheet is None:
        ctx = f"Sheet '{rest}' (details not found in report)."
    else:
        # Brief structural facts.
        # `_sheet_edges` keys are (src, tgt) where "src has a formula reading
        # from tgt". So sources-of-this-sheet = edges where src=this (its
        # formulas read from tgt), and consumers = edges where tgt=this
        # (other sheets read from us). Mirrors render._section_data_flow_story.
        edges = getattr(report, "_sheet_edges", None) or {}
        sources = sorted(
            ((tgt, cnt) for (src, tgt), cnt in edges.items() if src == sheet.name),
            key=lambda x: (-x[1], x[0]),
        )[:4]
        consumers = sorted(
            ((src, cnt) for (src, tgt), cnt in edges.items() if tgt == sheet.name),
            key=lambda x: (-x[1], x[0]),
        )[:4]
        srcs_str = ", ".join(f"{n}({c})" for n, c in sources) or "(none)"
        cons_str = ", ".join(f"{n}({c})" for n, c in consumers) or "(none)"
        ctx = (
            f"Sheet '{sheet.name}' ({sheet.state}, {sheet.rows_used}x"
            f"{sheet.cols_used}, {sheet.cells_nonempty} non-empty cells, "
            f"{sheet.cells_formula} formulas). "
            f"Sources (where its formulas read from): {srcs_str}. "
            f"Consumers (other sheets reading this one): {cons_str}."
        )
    return {
        "id": marker_id,
        "context": ctx,
        "ask": (
            "In 2-4 sentences, describe the business role of this sheet: "
            "what does it represent (e.g. 'master production schedule', "
            "'BOM lookup table', 'shift-level capacity input')? Is it user-"
            "edited input, computed output, or a transformation hub? Name "
            "the most likely business interpretation. Look at sample "
            "formulas in dossier.sheets[] to confirm."
        ),
    }


def _ctx_vba_narration(marker_id: str, rest: str, report) -> dict:
    narr = next((n for n in (report.vba_narratives or [])
                 if n.module_name == rest), None)
    cls = next((c for c in (report.vba_classifications or [])
                if c.module_name == rest), None)
    vm = next((v for v in (report.vba_modules or []) if v.name == rest), None)
    if narr is None and cls is None and vm is None:
        ctx = f"VBA module '{rest}' (no structured context available)."
    else:
        line_count = getattr(narr, "line_count", None) or getattr(vm, "line_count", 0)
        inferred = (getattr(narr, "inferred_type", None)
                    or getattr(cls, "inferred_type", None) or "mixed")
        reads = (getattr(narr, "reads_sheets", None)
                 or getattr(cls, "reads_sheets", None) or [])
        writes = (getattr(narr, "writes_sheets", None)
                  or getattr(cls, "writes_sheets", None) or [])
        callers = getattr(narr, "callers", None) or []
        callees = getattr(narr, "callees", None) or []
        oer = bool(getattr(narr, "has_oer", False)
                   or getattr(vm, "has_on_error_resume_next", False))
        external = bool(getattr(narr, "external_calls", False)
                        or (vm is not None and vm.external_keywords))
        ctx = (
            f"Module '{rest}' ({inferred}, {line_count} lines). "
            f"Reads sheets: {', '.join(reads[:4]) or '(none)'}. "
            f"Writes sheets: {', '.join(writes[:4]) or '(none)'}. "
            f"Called by: {', '.join(callers[:4]) or '(none)'}. "
            f"Calls: {', '.join(callees[:4]) or '(none)'}. "
            f"Has On Error Resume Next: {oer}. External/COM calls: {external}. "
            f"FULL source is in dossier.vba_modules[name=={rest!r}].source."
        )
    return {
        "id": marker_id,
        "context": ctx,
        "ask": (
            "Read this module's full source (in dossier.vba_modules) and "
            "answer in 2-4 sentences: (1) what is the module's BUSINESS "
            "purpose? (2) if you can name an academic or business method "
            "it implements (e.g. Holt-Winters smoothing, EOQ, MRP netting, "
            "VRP, Newsvendor, rolling-horizon scheduling), name it. "
            "(3) call out any reliability concerns: silent error swallow, "
            "race conditions, hardcoded magic, or fragile external calls."
        ),
    }


def _ctx_domain_method(marker_id: str, rest: str, report) -> dict:
    m = next((dt for dt in (report.domain_template_matches or [])
              if dt.template_key == rest), None)
    if m is None:
        ctx = f"Domain template '{rest}' (no match details in report)."
    else:
        kw = ", ".join(m.matched_keywords[:6]) or "(none)"
        methods = ", ".join(label for label, _ in (m.method_hits or [])[:5]) or "(none detected)"
        ctx = (
            f"Domain template '{m.business_friendly_name}' "
            f"(confidence: {m.confidence}). Matched keywords: {kw}. "
            f"Detected scheduling methods (heuristic): {methods}."
        )
    return {
        "id": marker_id,
        "context": ctx,
        "ask": (
            "In 2-4 sentences: (1) summarize how this domain template "
            "applies to this workbook (which sheet/VBA evidence supports "
            "it). (2) flag the highest-priority hardcode or method risk a "
            "domain expert should check first."
        ),
    }


# =============================================================================
# Dossier + prompt builders (public)
# =============================================================================

def build_dossier(report, audit_md_text: str, source_path=None) -> dict:
    """Build the structured context the LLM needs.

    Pure function: same inputs -> byte-identical output (after
    json.dumps(sort_keys=True)).

    `source_path`: path to the original xlsm; used to re-extract sample
    formulas per sheet. If omitted, sample_formulas will be [] for every
    sheet (graceful degradation — the LLM still has VBA source + structural
    summaries; sample formulas are nice-to-have).
    """
    cell_rows: list = []
    if source_path is not None:
        from .extract import extract_cells
        parse_errors: list = []
        try:
            cell_rows, _, _, _ = extract_cells(
                Path(source_path),
                parse_errors,
                sanitize=bool(report.sanitized),
            )
        except Exception:
            cell_rows = []

    markers = _extract_markers(audit_md_text)
    markers_to_fill = [_context_for_marker(mid, report) for mid in markers]

    return {
        "schema_version": "harness/1",
        "workbook_meta": _build_workbook_meta(report),
        "sheets": _build_sheets_section(report, cell_rows),
        "vba_modules": _build_vba_modules_section(report),
        "workflow": _build_workflow_section(report),
        "pillars_top10": _build_pillars_top(report),
        "smells_top10": _build_smells_top(report),
        "magic_numbers_top20": _build_magic_numbers_top(report),
        "anomalies_top10": _build_anomalies_top(report),
        "domain_templates": _build_domain_templates(report),
        "markers_to_fill": markers_to_fill,
    }


# =============================================================================
# Prompt construction
# =============================================================================

_PROMPT_HEADER = """\
# Excel Workbook Reverse-Engineering Task — Track B (BYOA Harness)

You are a senior reverse engineer of legacy Excel/VBA systems. A static-
analysis tool has produced a structural audit of one workbook and identified
specific spots where heuristic prose should be replaced with a SHORT, SHARP
business-semantic narrative written by you.

Your output will be ingested back into the audit by a deterministic substitution
step. You must follow the output format EXACTLY.
"""


# Glossary embedded into the prompt for DE/ZH consistency.
_PROMPT_GLOSSARY = """\
## Trilingual glossary — non-negotiable terminology

When producing German or Chinese narratives, use the canonical translations
below verbatim. When a term not in this table appears, prefer
SAP / manufacturing-industry-standard German or mainland Chinese conventions.
When uncertain, keep the English term in parentheses (e.g. "Stückliste (BOM)").

| English | German | Chinese |
|---|---|---|
| capacity planning | Kapazitätsplanung | 产能规划 |
| Bill of Materials | Stückliste | 物料清单 |
| MRP (Material Requirements Planning) | Materialbedarfsplanung | 物料需求计划 |
| MPS (Master Production Schedule) | Hauptproduktionsplan | 主生产计划 |
| pillar cell (high fan-in change point) | systemrelevante Zelle | 支柱单元格 |
| fan-in | Eingangsgrad / Referenzanzahl | 引用数(被引用次数) |
| smell | Code-Geruch | 代码异味 |
| anomaly | Anomalie | 异常 |
| heuristic | Heuristik | 启发式 |
| lead time | Lieferzeit / Vorlaufzeit | 提前期 |
| safety stock | Sicherheitsbestand | 安全库存 |
| hardcoded magic number | hartkodierte Magic Number | 硬编码魔法数字 |
| dead code (suspected) | toter Code (vermutet) | 可能的死代码 |
| workflow guide | Arbeitsablauf-Leitfaden | 工作流指南 |
| data flow | Datenfluss | 数据流 |
| executive summary | Zusammenfassung | 执行摘要 |
| complexity score | Komplexitätswert | 复杂度评分 |
| named range | benannter Bereich | 命名区域 |
| hidden / very-hidden sheet | ausgeblendetes / streng-ausgeblendetes Blatt | 隐藏 / 深度隐藏工作表 |
| static analysis | statische Analyse | 静态分析 |
| local-only / no network | nur lokal / keine Netzwerkverbindung | 仅本地 / 无网络调用 |
"""

_PROMPT_RULES = """\
## Rules for every narrative

1. **2-4 sentences each. No more.** This is selective replacement, not a deep
   dive. If you have more to say, the report has appendices for it.
2. **Business-semantic, not mechanical.**
   - YES: "Calculates net production requirements by netting forecast against
          on-hand inventory and open POs (classic MRP netting)."
   - NO:  "Loops over rows 2 through 200 and writes results to column G."
3. **Name methods when identifiable.** If the code implements a known
   academic / business method (Holt-Winters smoothing, EOQ, MRP netting,
   Newsvendor, VRP / vehicle-routing, rolling-horizon scheduling, capacity
   planning, ABC analysis, run-length encoding, SUMIFS-tile, VLOOKUP-tile,
   Bill of Materials traversal, etc.), say so explicitly. Don't hedge.
4. **Flag risk patterns.** When VBA does something risky — silent error
   swallow (`On Error Resume Next`), race-conditioned cell writes,
   hard-coded magic constants, fragile external/COM calls, fragile string
   manipulation — call it out in one clause.
5. **No preamble. No markdown. No fences.**
   Output is JSON ONLY (see "Output Format" below). Do not add commentary
   before or after the JSON.
6. **Don't echo the heuristic narrative.** The static analyzer has already
   written one — your job is to REPLACE it with semantic content.
7. **If you don't have enough info to write a confident narrative for an
   ID, OMIT IT** — the ingest step keeps the heuristic narrative when an ID
   is missing. Better to skip than to hallucinate.
"""

_PROMPT_OUTPUT_FORMAT = """\
## Output Format (STRICT)

Return **one** JSON object. Keys are the marker IDs listed in the "Questions"
section. Values are your narrative strings.

Example shape (illustrative — your real keys come from the Questions list):

```json
{
  "workflow-step:1": "Reloads forecast data from the corporate web service \
into the Forecast P1 sheet (this is the 'pull external numbers' step before \
MRP netting). The handler swallows errors silently — a network blip leaves \
stale forecast in place without warning.",
  "vba-narration:Module_MRP": "Implements classic MRP netting: nets gross \
requirements (from MPS) against on-hand inventory plus open supplier orders, \
producing planned order releases by part. Hardcodes a 0.05 loss factor that \
should arguably live in _constants.",
  "data-flow:BOM": "Bill of Materials lookup table. Static input — the \
parent/child component relationships are user-maintained and consumed by \
MRP netting downstream."
}
```

**Rules for the JSON object**:
- One top-level object. No outer array, no metadata wrapper.
- Keys are STRINGS exactly matching marker IDs in the Questions section.
- Values are STRINGS (one narrative each).
- Strings must be valid JSON (escape internal double-quotes and backslashes).
- Do NOT wrap the JSON in a markdown code fence. Plain JSON, top to bottom.
- Do NOT include keys you didn't get to. Empty narratives cause confusion;
  skip rather than emit "" or "TODO".
"""


_PROMPT_OUTPUT_FORMAT_TRILINGUAL = """\
## Output Format (STRICT, TRILINGUAL)

Return **one** JSON object. Keys are the marker IDs listed in the "Questions"
section. **Each value is itself a three-key object** with `en`, `de`, and
`zh` fields — one narrative per language.

Example shape (illustrative — your real keys come from the Questions list):

```json
{
  "vba-narration:Module_MRP": {
    "en": "Implements classic MRP netting: nets gross requirements (from MPS) \
against on-hand inventory plus open supplier orders, producing planned order \
releases by part. Hardcodes a 0.05 loss factor that should arguably live \
in _constants.",
    "de": "Implementiert klassisches MRP-Nettorechnen: verrechnet \
Bruttobedarfe (aus dem Hauptproduktionsplan) mit Lagerbestand plus offenen \
Lieferantenbestellungen und erzeugt geplante Bestellfreigaben pro Teil. \
Enthält einen hartkodierten Verlustfaktor 0.05, der eher in _constants \
gehören sollte.",
    "zh": "实现经典 MRP 净需求计算:用主生产计划的总需求扣减库存与未达供应商订单,\
按零件输出计划下单量。硬编码了 0.05 的损耗因子,更合适放在 _constants 中。"
  },
  "data-flow:BOM": {
    "en": "Bill of Materials lookup table. Static input — the parent/child \
component relationships are user-maintained and consumed by MRP netting \
downstream.",
    "de": "Stückliste (BOM) — Nachschlagetabelle. Statische Eingabe; die \
Eltern-Kind-Komponentenbeziehungen werden manuell gepflegt und nachgelagert \
vom MRP-Nettorechnen verwendet.",
    "zh": "物料清单查表。静态输入 — 父子组件关系由用户维护,下游被 MRP 净需求\
计算使用。"
  }
}
```

**Rules for the trilingual JSON object**:
- One top-level object. No outer array, no metadata wrapper.
- Keys are STRINGS exactly matching marker IDs in the Questions section.
- Values are OBJECTS with the three keys `en`, `de`, `zh`. All three fields
  required for every marker you fill.
- Each language field must contain 2-4 sentences with **the same factual
  content**; only the language should differ. Do NOT add facts to one language
  that the others don't have.
- Use the canonical glossary translations below — German favours SAP /
  manufacturing convention, Chinese favours mainland convention.
- Strings must be valid JSON (escape internal double-quotes and backslashes).
- Do NOT wrap the JSON in a markdown code fence. Plain JSON, top to bottom.
- If you cannot confidently produce all three languages for a marker, **omit
  that marker entirely** rather than mixing or hallucinating a translation.
  The ingest step keeps the heuristic narrative when an ID is missing.
"""


def _format_workbook_summary(meta: dict) -> str:
    """Compact 6-line summary for the prompt header."""
    parts = [
        f"- **Filename**: `{meta['filename']}`",
        f"- **Size**: {meta['file_size_bytes']:,} bytes; "
        f"SHA-256: `{meta['sha256'][:16]}...`",
        f"- **Sheets**: {meta['sheet_count']} total "
        f"({meta['sheet_count_visible']} visible, "
        f"{meta['sheet_count_hidden']} hidden, "
        f"{meta['sheet_count_very_hidden']} veryHidden)",
        f"- **Cells**: {meta['cell_count_nonempty']:,} non-empty "
        f"({meta['cell_count_formula']:,} formulas)",
        f"- **VBA**: {meta['vba_module_count']} modules / "
        f"{meta['vba_total_lines']:,} lines",
        f"- **Complexity score**: {meta['complexity_total']}/100",
    ]
    if meta.get("domain_detected"):
        parts.append(
            f"- **Domain (heuristic)**: {meta['domain_detected']} "
            f"(confidence: {meta['domain_confidence'] or 'n/a'})"
        )
    return "\n".join(parts)


def _format_questions(markers_to_fill: list) -> str:
    """Render the Questions section: one bullet per marker with context + ask."""
    if not markers_to_fill:
        return "_(No markers to fill — the report has nothing to augment.)_\n"
    lines: list = []
    for q in markers_to_fill:
        lines.append(f"### `{q['id']}`")
        lines.append("")
        lines.append(f"**Context**: {q['context']}")
        lines.append("")
        lines.append(f"**Ask**: {q['ask']}")
        lines.append("")
    return "\n".join(lines)


def build_prompt(dossier: dict, lang: str = "en") -> str:
    """Build the mega-prompt to give the LLM.

    The prompt embeds:
        - Header: role + task framing
        - Workbook summary (6-line compact)
        - Rules (2-4 sentences, semantic, name methods, flag risk, omit if unsure)
        - Output format (strict JSON, marker IDs as keys)
        - Questions: one entry per marker with context + ask

    `lang` selects the output schema:
        - "en"      — flat JSON (one narrative per marker)
        - "de"|"zh" — flat JSON, narratives in target language
        - "all"     — trilingual nested JSON ({en, de, zh} per marker)
                      with the glossary embedded.

    The dossier.json is referenced as an attachment — we tell the LLM
    "see dossier.json for full VBA source + sheet samples" so the prompt
    stays readable.
    """
    meta = dossier["workbook_meta"]
    questions = _format_questions(dossier["markers_to_fill"])
    n_markers = len(dossier["markers_to_fill"])

    is_trilingual = (lang == "all")
    if is_trilingual:
        output_format = _PROMPT_OUTPUT_FORMAT_TRILINGUAL
        glossary_block = "\n" + _PROMPT_GLOSSARY + "\n"
        closing = (
            "Remember: respond with ONE JSON object, no markdown fence, "
            "no preamble. Keys = marker IDs above; values = three-key objects "
            "with `en`, `de`, `zh` narratives (2-4 sentences each, same facts, "
            "only the language differs). Skip any ID you can't confidently "
            "answer in all three languages.\n"
        )
    elif lang in ("de", "zh"):
        # Single-lang non-English: reuse flat shape but instruct narratives
        # to be in target language using glossary.
        lang_word = {"de": "German (SAP/manufacturing convention)",
                     "zh": "Chinese (mainland convention)"}[lang]
        output_format = (
            _PROMPT_OUTPUT_FORMAT
            + f"\n**Language**: produce all narratives in **{lang_word}**, "
            f"using the canonical glossary below verbatim. Keep technical "
            f"identifiers (sheet names, cell refs, VBA module names) "
            f"unchanged in any language.\n"
        )
        glossary_block = "\n" + _PROMPT_GLOSSARY + "\n"
        closing = (
            f"Remember: respond with ONE JSON object, no markdown fence, no "
            f"preamble. Keys = marker IDs above; values = your 2-4 sentence "
            f"narratives in {lang_word}. Skip any ID you can't confidently "
            f"answer.\n"
        )
    else:
        output_format = _PROMPT_OUTPUT_FORMAT
        glossary_block = ""
        closing = (
            "Remember: respond with ONE JSON object, no markdown fence, no "
            "preamble. Keys = marker IDs above; values = your 2-4 sentence "
            "narratives. Skip any ID you can't confidently answer.\n"
        )

    body = (
        _PROMPT_HEADER
        + "\n## Workbook context\n\n"
        + _format_workbook_summary(meta)
        + "\n\n"
        + _PROMPT_RULES
        + "\n"
        + output_format
        + glossary_block
        + (
            f"## Reference data\n\n"
            f"The companion file **`dossier.json`** (attached alongside this "
            f"prompt) contains:\n"
            f"- Full VBA source for every module (`vba_modules[].source`)\n"
            f"- Up to {_SAMPLE_FORMULAS_PER_SHEET} sample formulas per sheet "
            f"(`sheets[].sample_formulas`)\n"
            f"- Top pillars / smells / magic numbers / anomalies / domain "
            f"templates\n"
            f"- The complete list of markers to fill (`markers_to_fill[]`) "
            f"with structured context for each\n\n"
            f"You should attach (drag in) `dossier.json` along with this "
            f"prompt before answering. The prompt below restates each "
            f"question's context inline so you don't have to cross-reference "
            f"if your tool can't see attachments.\n\n"
        )
        + f"## Questions ({n_markers} markers)\n\n"
        + questions
        + "\n---\n\n"
        + closing
    )
    return body


# =============================================================================
# Public entry point
# =============================================================================

def extract(report, audit_md_text: str, out_dir: Path,
            source_path=None, lang: str = "en") -> dict:
    """Extract phase — write dossier.json + prompt.md into out_dir.

    Returns paths dict for the CLI to print/log.

    `audit_md_text` is the rendered audit.md content (so we can pull marker
    IDs the renderer actually emitted). `out_dir` must exist. `source_path`
    is the original xlsm path; passing it enables sample-formula extraction
    per sheet. `lang` selects the prompt schema (`en`|`de`|`zh`|`all`); when
    `all`, the prompt requests the trilingual nested JSON shape.
    """
    out_dir = Path(out_dir)
    if not out_dir.exists():
        raise RuntimeError(f"out_dir does not exist: {out_dir}")

    dossier = build_dossier(report, audit_md_text, source_path=source_path)
    # Echo the requested language into the dossier so ingest can sanity-check
    # whether the responses.json shape it received matches what the prompt
    # asked for.
    dossier["prompt_lang"] = lang
    prompt = build_prompt(dossier, lang=lang)

    dossier_path = out_dir / "dossier.json"
    prompt_path = out_dir / "prompt.md"

    # sort_keys + UTF-8: deterministic output for idempotency tests
    dossier_path.write_text(
        json.dumps(dossier, sort_keys=True, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    prompt_path.write_text(prompt, encoding="utf-8")

    return {
        "dossier_path": str(dossier_path),
        "prompt_path": str(prompt_path),
        "marker_count": len(dossier["markers_to_fill"]),
        "lang": lang,
    }
