"""Hermans-style smell detectors.

Six detectors:
    multiple-references / long-calculation-chain / conditional-complexity /
    multiple-operations / magic-numbers / duplicated-formulas

Each returns SmellFinding lists. The internal maps (`incoming` for fan-in,
`pattern_to_cells` for duplicated-formula clusters) are exported as side
products so the new logic-comprehension modules (pillars, anomalies) can
consume them without re-tokenizing the whole workbook.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass

from .formula_utils import (
    _conditional_nesting_depth,
    _expand_range_to_cells,
    _extract_ranges,
    _normalize_pattern,
    _operator_count,
    _token_attr,
    _token_class,
    _tokenize,
)

# =============================================================================
# Thresholds (calibrated from Stage 1 corpus)
# =============================================================================

THRESH_MULTIPLE_REFS = 5
THRESH_LONG_CHAIN = 5
THRESH_CONDITIONAL_NESTING = 2
THRESH_MULTIPLE_OPS = 8
THRESH_DUPLICATED_PATTERN = 5

MAX_CHAIN_DEPTH_CAP = 50
MAX_FORMULAS_FOR_CHAIN = 20000

TRIVIAL_NUMBERS = frozenset({"0", "1", "-1", "2", "100", "1.0", "0.0", "10", "0.5"})

SMELL_TYPES = [
    "multiple-references",
    "long-calculation-chain",
    "conditional-complexity",
    "multiple-operations",
    "magic-numbers",
    "duplicated-formulas",
]


@dataclass
class SmellFinding:
    smell_type: str
    location: str
    severity: str
    confidence: str
    evidence: str
    metric: float


@dataclass
class MagicNumberEntry:
    value: str
    occurrence_count: int
    first_location: str
    sample_context: str
    location_kind: str
    confidence: str


# =============================================================================
# Detectors
# =============================================================================

def detect_multiple_references(
    cell_rows: list, sheet_names: set,
) -> tuple[list, dict, dict]:
    """
    Returns (findings, outgoing_edges, incoming_edges).

    `incoming` maps target 'sheet|REF' -> set of source 'sheet|ref' formulas.
    Pillar-cell analysis (A.1) reuses `incoming` directly.
    """
    incoming: dict = defaultdict(set)
    sheet_names_lower = {s.lower(): s for s in sheet_names}
    for cr in cell_rows:
        if not cr.formula:
            continue
        tokens, ok = _tokenize(cr.formula)
        if not ok:
            continue
        ranges = _extract_ranges(tokens, default_sheet=cr.sheet, sheet_names=sheet_names)
        for r in ranges:
            sh = sheet_names_lower.get(r["sheet"].lower(), r["sheet"]) if r["sheet"] else cr.sheet
            for cell_ref in _expand_range_to_cells(r["ref"], sh):
                key = f"{sh}|{cell_ref}"
                incoming[key].add(f"{cr.sheet}|{cr.ref}")

    findings: list = []
    for key, srcs in incoming.items():
        if len(srcs) < THRESH_MULTIPLE_REFS:
            continue
        sheet, ref = key.split("|", 1)
        findings.append(SmellFinding(
            smell_type="multiple-references",
            location=f"{sheet}!{ref}",
            severity="high" if len(srcs) >= 20 else ("medium" if len(srcs) >= 10 else "low"),
            confidence="high",
            evidence=f"referenced by {len(srcs)} distinct formulas",
            metric=float(len(srcs)),
        ))
    findings.sort(key=lambda f: (-f.metric, f.location))
    findings = findings[:50]

    outgoing: dict = defaultdict(set)
    for tgt, srcs in incoming.items():
        for src in srcs:
            outgoing[src].add(tgt)
    return findings, outgoing, incoming


def detect_long_calculation_chains(cell_rows: list, outgoing: dict) -> list:
    if len(cell_rows) > MAX_FORMULAS_FOR_CHAIN:
        return [SmellFinding(
            smell_type="long-calculation-chain",
            location="<analysis-skipped>",
            severity="low", confidence="low",
            evidence=f"workbook exceeds {MAX_FORMULAS_FOR_CHAIN} formula cells; chain analysis skipped",
            metric=0.0,
        )]

    formula_cells = {f"{cr.sheet}|{cr.ref}" for cr in cell_rows if cr.formula}
    depth: dict = {}
    state: dict = {}

    def dfs(node: str) -> int:
        if state.get(node) == 2:
            return depth.get(node, 0)
        if state.get(node) == 1:
            return 0
        state[node] = 1
        max_child = 0
        for child in outgoing.get(node, ()):
            if child in formula_cells:
                d = dfs(child)
                if d > max_child:
                    max_child = d
                if max_child >= MAX_CHAIN_DEPTH_CAP:
                    break
        depth[node] = 1 + max_child
        state[node] = 2
        return depth[node]

    findings: list = []
    for node in formula_cells:
        if state.get(node) != 2:
            try:
                dfs(node)
            except RecursionError:
                continue

    for node, d in depth.items():
        if d < THRESH_LONG_CHAIN:
            continue
        sheet, ref = node.split("|", 1)
        findings.append(SmellFinding(
            smell_type="long-calculation-chain",
            location=f"{sheet}!{ref}",
            severity="high" if d >= 10 else ("medium" if d >= 7 else "low"),
            confidence="medium",
            evidence=f"transitive dependency depth {d}",
            metric=float(d),
        ))
    findings.sort(key=lambda f: (-f.metric, f.location))
    return findings[:30]


def detect_conditional_complexity(cell_rows: list) -> list:
    findings: list = []
    for cr in cell_rows:
        if not cr.formula:
            continue
        tokens, ok = _tokenize(cr.formula)
        if not ok:
            continue
        d = _conditional_nesting_depth(tokens)
        if d < THRESH_CONDITIONAL_NESTING:
            continue
        findings.append(SmellFinding(
            smell_type="conditional-complexity",
            location=f"{cr.sheet}!{cr.ref}",
            severity="high" if d >= 5 else "medium",
            confidence="medium",
            evidence=f"IF/IFS/CHOOSE nesting depth {d}",
            metric=float(d),
        ))
    findings.sort(key=lambda f: (-f.metric, f.location))
    return findings[:30]


def detect_multiple_operations(cell_rows: list) -> list:
    findings: list = []
    for cr in cell_rows:
        if not cr.formula:
            continue
        tokens, ok = _tokenize(cr.formula)
        if not ok:
            continue
        n = _operator_count(tokens)
        if n < THRESH_MULTIPLE_OPS:
            continue
        findings.append(SmellFinding(
            smell_type="multiple-operations",
            location=f"{cr.sheet}!{cr.ref}",
            severity="high" if n >= 20 else ("medium" if n >= 12 else "low"),
            confidence="medium",
            evidence=f"{n} operators+function-calls in one formula",
            metric=float(n),
        ))
    findings.sort(key=lambda f: (-f.metric, f.location))
    return findings[:30]


def detect_duplicated_formulas(cell_rows: list) -> tuple[list, dict]:
    """
    Returns (findings, pattern_to_cells).

    `pattern_to_cells` maps normalized-pattern string -> list of CellRow refs
    (the actual cells that share each pattern). A.2 anomaly detection reuses
    this map to avoid a second pass.
    """
    pattern_counts: Counter = Counter()
    pattern_examples: dict = {}
    pattern_to_cells: dict = defaultdict(list)
    for cr in cell_rows:
        if not cr.formula:
            continue
        tokens, ok = _tokenize(cr.formula)
        if not ok:
            continue
        pat = _normalize_pattern(tokens)
        if not pat:
            continue
        pattern_counts[pat] += 1
        pattern_to_cells[pat].append(cr)
        if pat not in pattern_examples:
            pattern_examples[pat] = (f"{cr.sheet}!{cr.ref}", cr.formula)

    findings: list = []
    for pat, count in pattern_counts.items():
        if count < THRESH_DUPLICATED_PATTERN:
            continue
        loc, sample = pattern_examples[pat]
        findings.append(SmellFinding(
            smell_type="duplicated-formulas",
            location=f"pattern@{loc}",
            severity="high" if count >= 50 else ("medium" if count >= 20 else "low"),
            confidence="medium",
            evidence=f"{count} cells share this normalized formula pattern; sample: {sample[:80]}",
            metric=float(count),
        ))
    findings.sort(key=lambda f: (-f.metric, f.location))
    return findings[:20], pattern_to_cells


def _looks_like_string(line: str) -> bool:
    in_str = False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return False
    return in_str


def detect_magic_numbers(cell_rows: list, vba_text: str) -> tuple[list, list]:
    counter: Counter = Counter()
    samples: dict = {}
    samples_kind: dict = {}
    samples_loc: dict = {}

    cell_with_magic: dict = defaultdict(list)
    for cr in cell_rows:
        if not cr.formula:
            continue
        tokens, ok = _tokenize(cr.formula)
        if not ok:
            for m in re.finditer(r"(?<![\w\.\$])(\d+\.?\d*)(?![\w\.])", cr.formula):
                v = m.group(1)
                if v in TRIVIAL_NUMBERS:
                    continue
                counter[v] += 1
                cell_with_magic[f"{cr.sheet}|{cr.ref}"].append(v)
                if v not in samples:
                    samples[v] = cr.formula[:120]
                    samples_kind[v] = "cell"
                    samples_loc[v] = f"{cr.sheet}!{cr.ref}"
            continue
        for tok in tokens:
            if _token_class(tok) != "Number":
                continue
            attr = _token_attr(tok)
            v = str(attr.get("name", "")).strip()
            if not v or v in TRIVIAL_NUMBERS:
                continue
            if v.upper() in ("TRUE", "FALSE"):
                continue
            if not any(ch.isdigit() for ch in v):
                continue
            counter[v] += 1
            cell_with_magic[f"{cr.sheet}|{cr.ref}"].append(v)
            if v not in samples:
                samples[v] = cr.formula[:120]
                samples_kind[v] = "cell"
                samples_loc[v] = f"{cr.sheet}!{cr.ref}"

    _GUID_LIKE = re.compile(r"\{[0-9A-Fa-f\-]{8,}\}|VB_Base|VB_GlobalNameSpace|VB_TemplateDerived|VB_Customizable")
    _SEP_LIKE = re.compile(r"^=+$")
    for line in vba_text.splitlines():
        if _GUID_LIKE.search(line):
            continue
        if _SEP_LIKE.match(line.strip()):
            continue
        if line.strip().startswith("Attribute "):
            continue
        if line.startswith("=== Module:") or line.startswith("=== Stream:"):
            continue
        line_nocomment = line.split("'", 1)[0] if "'" in line and not _looks_like_string(line) else line
        for m in re.finditer(r"(?<![\w\.])(-?\d+\.?\d*)(?![\w\.])", line_nocomment):
            v = m.group(1)
            if v.lstrip("-") in TRIVIAL_NUMBERS:
                continue
            if len(v) >= 6 and v.startswith("0") and "." not in v:
                continue
            counter[v] += 1
            if v not in samples:
                samples[v] = line.strip()[:120]
                samples_kind[v] = "vba"
                samples_loc[v] = "<vba>"

    items = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))[:20]
    index: list = []
    for v, c in items:
        index.append(MagicNumberEntry(
            value=v,
            occurrence_count=c,
            first_location=samples_loc.get(v, ""),
            sample_context=samples.get(v, ""),
            location_kind=samples_kind.get(v, "cell"),
            confidence="medium",
        ))

    findings: list = []
    for key, vals in cell_with_magic.items():
        sheet, ref = key.split("|", 1)
        findings.append(SmellFinding(
            smell_type="magic-numbers",
            location=f"{sheet}!{ref}",
            severity="high" if len(vals) >= 5 else ("medium" if len(vals) >= 2 else "low"),
            confidence="medium",
            evidence=f"{len(vals)} non-trivial numeric literal(s): {','.join(sorted(set(vals))[:6])}",
            metric=float(len(vals)),
        ))
    findings.sort(key=lambda f: (-f.metric, f.location))
    findings = findings[:50]
    return findings, index
