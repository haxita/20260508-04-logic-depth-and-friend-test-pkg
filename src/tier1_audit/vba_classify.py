"""VBA module purpose inference (A.3) — heuristic structural classification.

Pure rule-based, zero LLM. We don't claim to know what a module *does*
semantically — only the *shape* of what it does:

    data-loader      — names like Load*/Get*/Read*/Import*; many .Value reads
    transformer      — names like Calc*/Compute*/Process*/Update*; balanced reads/writes
    report-writer    — names like Print*/Export*/Output*/Write*/Save*; many .Value writes
    ui-handler       — names like *_Click/*_Change/*_Activate/Workbook_Open; UI events
    dead-suspected   — no Sub/Function calls anywhere AND no .Value writes
    mixed            — fallback when no single signal dominates

We also extract:
    - reads_sheets, writes_sheets — referenced sheet names
    - external_calls — boolean (already in Stage 1's external_keywords)
    - control_flow_count — sum of For/Do While/Select Case/If counts
    - value_writes / value_reads — number of `Range(...).Value =` assignments vs reads
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

# Name-pattern → category, in priority order.
# A module's category is decided by which pattern wins the popularity vote
# across its Sub/Function names.
_NAME_PATTERNS = [
    ("data-loader", re.compile(r"^(?:Load|Get|Read|Fetch|Import|Pull|Refresh|Open)\w*", re.IGNORECASE)),
    ("transformer", re.compile(r"^(?:Calc|Compute|Process|Update|Build|Generate|Make|Run|Execute|Solve|Apply|Transform)\w*", re.IGNORECASE)),
    ("report-writer", re.compile(r"^(?:Print|Export|Output|Write|Save|Send|Email|Publish|Report|Dump)\w*", re.IGNORECASE)),
    ("ui-handler", re.compile(r"(?:_Click|_Change|_Activate|_Deactivate|_BeforeClose|_Open|_BeforeSave|_SelectionChange|_DblClick)$", re.IGNORECASE)),
]

# UI-handler module names from openpyxl/oletools: "ThisWorkbook", "Sheet1" etc.
_UI_MODULE_NAMES = re.compile(r"^(?:ThisWorkbook|Sheet\d+|Workbook|Worksheet)", re.IGNORECASE)

_RE_VALUE_WRITE = re.compile(
    r"""(?ix)
    \b(?:Range|Cells)\s*\([^)]*\)\s*(?:\.\s*Value)?\s*=
    """
)
_RE_VALUE_READ = re.compile(
    r"""(?ix)
    =\s*[^=]*\b(?:Range|Cells)\s*\([^)]*\)\s*\.\s*Value\b
    """
)
_RE_SHEET_REF = re.compile(
    r"""(?ix)
    \b(?:Sheets|Worksheets)\s*\(\s*
    "(?P<name>[^"]+)"
    \s*\)
    """
)
_RE_FOR = re.compile(r"(?im)^\s*For\b")
_RE_DO = re.compile(r"(?im)^\s*Do\b")
_RE_WHILE = re.compile(r"(?im)^\s*While\b")
_RE_SELECT_CASE = re.compile(r"(?im)^\s*Select\s+Case\b")
_RE_IF = re.compile(r"(?im)^\s*If\b")

# `Call <name>` and bare `<name> arg1, arg2` invocations.
_RE_CALL = re.compile(r"(?im)^\s*Call\s+([A-Za-z_]\w*)")


@dataclass
class VbaClassification:
    module_name: str
    inferred_type: str          # one of {data-loader, transformer, report-writer, ui-handler, dead-suspected, mixed}
    confidence: str             # 'high' | 'medium' | 'low'
    reads_sheets: list          # sorted unique
    writes_sheets: list         # sorted unique (best-effort: a sheet referenced in a write context)
    external_calls: bool        # any external/COM keyword present
    value_writes: int           # count of Range/Cells(...).Value = ... patterns
    value_reads: int            # count of `... = ...Range/Cells(...).Value` patterns
    control_flow_count: int     # For + Do/While + Select Case + If
    name_signals: dict          # category -> count of matching Sub/Function names (deterministic ordering preserved by str-key sort)
    rationale: str              # one-line human-readable explanation


def _scan_text(code: str) -> dict:
    """Compute base counts that the classifier needs."""
    return {
        "value_writes": len(list(_RE_VALUE_WRITE.finditer(code))),
        "value_reads": len(list(_RE_VALUE_READ.finditer(code))),
        "for": len(list(_RE_FOR.finditer(code))),
        "do": len(list(_RE_DO.finditer(code))),
        "while": len(list(_RE_WHILE.finditer(code))),
        "select_case": len(list(_RE_SELECT_CASE.finditer(code))),
        "if": len(list(_RE_IF.finditer(code))),
        "calls": len(list(_RE_CALL.finditer(code))),
    }


def _extract_sheet_refs(code: str) -> list:
    """Return sorted unique sheet names referenced via Sheets("X")/Worksheets("X")."""
    return sorted({m.group("name") for m in _RE_SHEET_REF.finditer(code)})


def _classify_one_module(module) -> VbaClassification:
    code = getattr(module, "source_text", "") or ""
    counts = _scan_text(code)

    # Tally name-pattern signals from Sub/Function names
    name_signals: Counter = Counter()
    for sf in module.sub_functions:
        for cat, pat in _NAME_PATTERNS:
            if pat.search(sf.name):
                name_signals[cat] += 1
                break  # one match per name

    # If the module name itself is a UI module (ThisWorkbook, SheetN), bias toward ui-handler
    if _UI_MODULE_NAMES.search(module.name) and any(
        sf.kind != "Function" for sf in module.sub_functions
    ):
        # but only if there are event-style Sub names too
        if any(re.search(r"(?:_Click|_Change|_Activate|_Open|_BeforeClose|_BeforeSave|_SelectionChange)$",
                         sf.name, re.IGNORECASE)
               for sf in module.sub_functions):
            name_signals["ui-handler"] += 5  # strong bias

    sheet_refs = _extract_sheet_refs(code)

    # writes_sheets: sheets referenced within ±200 chars of a Value-write match.
    writes_sheets: set = set()
    reads_sheets: set = set()
    for m in _RE_VALUE_WRITE.finditer(code):
        ctx = code[max(0, m.start() - 200): m.end() + 50]
        for sm in _RE_SHEET_REF.finditer(ctx):
            writes_sheets.add(sm.group("name"))
    for m in _RE_VALUE_READ.finditer(code):
        ctx = code[max(0, m.start() - 200): m.end() + 50]
        for sm in _RE_SHEET_REF.finditer(ctx):
            reads_sheets.add(sm.group("name"))
    # Sheets referenced but not classified as read/write get attributed to reads
    # (conservative — assume read access).
    only_referenced = set(sheet_refs) - writes_sheets - reads_sheets
    reads_sheets |= only_referenced

    external = bool(module.external_keywords)
    cf_count = counts["for"] + counts["do"] + counts["while"] + counts["select_case"] + counts["if"]

    # ---------- decision ----------
    # Step 1: dead-suspected — empty module shell with no executable content.
    # We treat Sheet*.cls / ThisWorkbook.cls with 0 subs as dead too: openpyxl/oletools
    # ship a .cls per worksheet by default, even if the workbook has no event handlers.
    is_empty_shell = (
        len(module.sub_functions) == 0
        and counts["calls"] == 0
        and counts["value_writes"] == 0
        and counts["value_reads"] == 0
        and module.line_count < 30  # default attribute headers are ~8 lines
    )
    no_user_code = (
        counts["calls"] == 0 and counts["value_writes"] == 0
        and not name_signals and module.line_count < 80
    )
    if is_empty_shell or (no_user_code and not _UI_MODULE_NAMES.search(module.name)):
        inferred_type = "dead-suspected"
        confidence = "medium"
        rationale = (f"empty/near-empty shell: {len(module.sub_functions)} subs, "
                     f"{counts['calls']} calls, {counts['value_writes']} writes, "
                     f"{module.line_count} LOC")
    else:
        # Step 2: name signal vote
        if name_signals:
            top_cat, top_n = name_signals.most_common(1)[0]
            second_n = name_signals.most_common(2)[1][1] if len(name_signals) > 1 else 0
            margin = top_n - second_n
            inferred_type = top_cat
            if margin >= 2:
                confidence = "high"
            elif top_n >= 2:
                confidence = "medium"
            else:
                confidence = "low"
            rationale = (f"name pattern vote: {top_cat}={top_n}"
                         + (f", others={dict(sorted(name_signals.items()))}" if len(name_signals) > 1 else ""))
        # Step 3: structural fallback when names are uninformative
        elif counts["value_writes"] >= 5 and counts["value_writes"] >= counts["value_reads"] * 2:
            inferred_type = "report-writer"
            confidence = "low"
            rationale = (f"name uninformative; {counts['value_writes']} .Value writes "
                         f"vs {counts['value_reads']} reads")
        elif counts["value_reads"] >= 5 and counts["value_reads"] >= counts["value_writes"] * 2:
            inferred_type = "data-loader"
            confidence = "low"
            rationale = (f"name uninformative; {counts['value_reads']} .Value reads "
                         f"vs {counts['value_writes']} writes")
        # Transformer: meaningful loops + value reads AND value writes
        elif (counts["value_reads"] >= 3 and counts["value_writes"] >= 3
              and (counts["for"] + counts["do"] + counts["while"]) >= 1):
            inferred_type = "transformer"
            confidence = "low"
            rationale = (f"name uninformative; reads={counts['value_reads']}, "
                         f"writes={counts['value_writes']}, loops="
                         f"{counts['for'] + counts['do'] + counts['while']}")
        elif cf_count >= 5 and counts["calls"] >= 2:
            inferred_type = "transformer"
            confidence = "low"
            rationale = f"name uninformative; control-flow={cf_count}, calls={counts['calls']}"
        else:
            inferred_type = "mixed"
            confidence = "low"
            rationale = "no decisive signal — name-pattern empty and structural counts mixed"

    return VbaClassification(
        module_name=module.name,
        inferred_type=inferred_type,
        confidence=confidence,
        reads_sheets=sorted(reads_sheets),
        writes_sheets=sorted(writes_sheets),
        external_calls=external,
        value_writes=counts["value_writes"],
        value_reads=counts["value_reads"],
        control_flow_count=cf_count,
        name_signals=dict(sorted(name_signals.items())),
        rationale=rationale,
    )


def classify_modules(vba_modules: list) -> list:
    """Classify every VBA module. Sorted by module_name (case-insensitive)."""
    out = [_classify_one_module(m) for m in vba_modules]
    out.sort(key=lambda c: c.module_name.lower())
    return out
