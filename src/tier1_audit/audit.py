"""Audit orchestrator.

Glues together extract / smells / pillars / anomalies / vba_classify into a
single AuditReport object that the renderers can serialize to markdown + json.

Design note: we keep the orchestration here (not in cli.py) so callers from
inside Python (tests, future Tier 1.5 hooks) get the same data path the CLI uses.
"""

from __future__ import annotations

import hashlib
import math
import re
import warnings
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

warnings.filterwarnings("ignore", message="Unknown extension is not supported")
warnings.filterwarnings("ignore", message="Workbook contains no default style")
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import openpyxl  # noqa: E402

from .anomalies import detect_anomalies
from .domain import detect_domain
from .domains import evaluate_domain_templates
from .extract import extract_cells, extract_vba, VBA_EXTERNAL_KEYWORDS
from .formula_utils import _extract_ranges, _tokenize
from .pillars import detect_pillars
from .smells import (
    SMELL_TYPES,
    THRESH_CONDITIONAL_NESTING,
    THRESH_DUPLICATED_PATTERN,
    THRESH_LONG_CHAIN,
    THRESH_MULTIPLE_OPS,
    THRESH_MULTIPLE_REFS,
    TRIVIAL_NUMBERS,
    detect_conditional_complexity,
    detect_duplicated_formulas,
    detect_long_calculation_chains,
    detect_magic_numbers,
    detect_multiple_operations,
    detect_multiple_references,
)
from .vba_classify import classify_modules
from .vba_narrate import narrate_modules
from .workflow import detect_workflow

# Capture library versions for the methodology footer (deterministic)
try:
    import oletools as _oletools_module  # noqa
    _OLETOOLS_VERSION = getattr(_oletools_module, "__version__", "unknown")
except Exception:
    _OLETOOLS_VERSION = "unknown"
try:
    import formulas as _formulas_module  # noqa
    _FORMULAS_VERSION = getattr(_formulas_module, "__version__", "unknown")
except Exception:
    _FORMULAS_VERSION = "unknown"

FORMULA_ERROR_TOKENS = ("#REF!", "#NAME?", "#DIV/0!", "#VALUE!", "#N/A", "#NUM!", "#NULL!")


# =============================================================================
# Dataclasses
# =============================================================================

@dataclass
class FileMeta:
    file_name: str
    file_size_bytes: int
    sha256: str


@dataclass
class BasicStats:
    sheet_count: int
    sheet_count_visible: int
    sheet_count_hidden: int
    sheet_count_very_hidden: int
    cell_count_nonempty: int
    cell_count_formula: int
    cell_count_unique_values: int
    named_range_count: int
    conditional_formatting_count: int
    data_validation_count: int
    vba_module_count: int
    vba_total_lines: int
    parse_errors_count: int


@dataclass
class SheetSummary:
    name: str
    state: str
    rows_used: int
    cols_used: int
    cells_nonempty: int
    cells_formula: int
    max_ref: str
    conditional_formatting_count: int
    data_validation_count: int


@dataclass
class RiskIndicators:
    hidden_sheets: list
    very_hidden_sheets: list
    cross_sheet_reference_count: int
    cells_with_errors: list
    external_workbook_references: list
    circular_reference_suspects: list
    parse_errors: list


@dataclass
class ComplexitySubScores:
    data_scale: int
    formula_depth: int
    vba_mass: int
    smell_density: int
    metadata_complexity: int


@dataclass
class ComplexityScore:
    total: int
    sub_scores: ComplexitySubScores
    rationale: dict


@dataclass
class AuditReport:
    meta: FileMeta
    basic_stats: BasicStats
    sheets: list
    named_ranges: list
    complexity: ComplexityScore
    smells: list
    magic_numbers: list
    vba_modules: list
    vba_classifications: list  # NEW
    pillars: list              # NEW (A.1)
    anomalies: list            # NEW (A.2)
    risk_indicators: RiskIndicators
    methodology: dict
    sanitized: bool
    # Polish-round additions:
    domain_hint: object = None       # tier1_audit.domain.DomainHint
    # Round-3 additions:
    workflow: dict = None            # output of workflow.detect_workflow()
    vba_narratives: list = None      # list[VbaNarrative]
    domain_template_matches: list = None  # list[DomainTemplateMatch]
    # _sheet_edges is a private analytical product used only by the renderer
    # for the Mermaid sheet-flow diagram. It's keyed by (src_sheet, tgt_sheet)
    # tuple -> int count. Excluded from JSON serialization (leading underscore
    # convention; see render._to_jsonable).
    _sheet_edges: dict = None


# =============================================================================
# Helpers
# =============================================================================

def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


_RE_EXTERNAL_REF = re.compile(r"\[[^\[\]]+\.(xls[xmb]?|xlsm)\]", re.IGNORECASE)


def _build_risk_indicators(
    cell_rows: list, sheet_meta: dict, parse_errors: list, outgoing: dict,
) -> RiskIndicators:
    hidden = sorted([n for n, m in sheet_meta.items() if m["state"] == "hidden"], key=str.lower)
    very_hidden = sorted([n for n, m in sheet_meta.items() if m["state"] == "veryHidden"], key=str.lower)

    cross_count = 0
    error_cells: list = []
    external_refs: set = set()
    sheet_names = set(sheet_meta.keys())

    for cr in cell_rows:
        if cr.formula:
            tokens, ok = _tokenize(cr.formula)
            if ok:
                ranges = _extract_ranges(tokens, default_sheet=cr.sheet, sheet_names=sheet_names)
                cited = {r["sheet"] for r in ranges if r["is_cross_sheet"]}
                if cited:
                    cross_count += 1
            for m in _RE_EXTERNAL_REF.finditer(cr.formula):
                external_refs.add(m.group(0))
        if cr.cached_value:
            for tok in FORMULA_ERROR_TOKENS:
                if tok in cr.cached_value:
                    error_cells.append({"sheet": cr.sheet, "ref": cr.ref, "error_token": tok})
                    break
        elif cr.value:
            for tok in FORMULA_ERROR_TOKENS:
                if tok in cr.value:
                    error_cells.append({"sheet": cr.sheet, "ref": cr.ref, "error_token": tok})
                    break

    error_cells.sort(key=lambda e: (e["sheet"], e["ref"]))
    external_refs_list = sorted(external_refs)

    in_cycle: set = set()
    visited: dict = {}
    for start in list(outgoing.keys()):
        if visited.get(start, 0) != 0:
            continue
        stack = [(start, iter(outgoing.get(start, ())))]
        path = [start]
        path_set = {start}
        visited[start] = 1
        while stack:
            node, it = stack[-1]
            advanced = False
            for nxt in it:
                if visited.get(nxt, 0) == 0:
                    visited[nxt] = 1
                    path.append(nxt)
                    path_set.add(nxt)
                    stack.append((nxt, iter(outgoing.get(nxt, ()))))
                    advanced = True
                    break
                elif visited.get(nxt, 0) == 1:
                    if nxt in path_set:
                        idx = path.index(nxt)
                        for n in path[idx:]:
                            in_cycle.add(n)
            if not advanced:
                done_node, _ = stack.pop()
                visited[done_node] = 2
                if path and path[-1] == done_node:
                    path.pop()
                    path_set.discard(done_node)

    circ_fmt = sorted({f"{n.split('|',1)[0]}!{n.split('|',1)[1]}" for n in in_cycle})

    parse_errors_sorted = sorted(parse_errors, key=lambda e: (str(e.get("sheet", "")),
                                                              str(e.get("ref", "")),
                                                              str(e.get("error", ""))))

    return RiskIndicators(
        hidden_sheets=hidden,
        very_hidden_sheets=very_hidden,
        cross_sheet_reference_count=cross_count,
        cells_with_errors=error_cells[:200],
        external_workbook_references=external_refs_list,
        circular_reference_suspects=circ_fmt[:100],
        parse_errors=parse_errors_sorted[:200],
    )


def _compute_complexity(
    basic: BasicStats, sheets: list, smells: list, risk: RiskIndicators,
) -> ComplexityScore:
    n_cells = max(1, basic.cell_count_nonempty)
    ds = min(20, int(math.log10(n_cells) * 6))

    cond_smells = [s for s in smells if s.smell_type == "conditional-complexity"]
    multi_op_smells = [s for s in smells if s.smell_type == "multiple-operations"]
    chain_smells = [s for s in smells if s.smell_type == "long-calculation-chain"
                    and s.location != "<analysis-skipped>"]
    max_cond = max((s.metric for s in cond_smells), default=0)
    max_op = max((s.metric for s in multi_op_smells), default=0)
    max_chain = max((s.metric for s in chain_smells), default=0)
    fd = min(20, int(max_cond * 2 + max_op * 0.4 + max_chain * 0.8))

    vm = min(20, int(math.log10(max(basic.vba_total_lines, 1)) * 4 + min(basic.vba_module_count, 10) * 0.5))

    sd_raw = (len(smells) / max(basic.cell_count_nonempty, 1)) * 1000
    sd = min(20, int(sd_raw * 4))

    metadata_score = (
        len(risk.hidden_sheets) * 2 +
        len(risk.very_hidden_sheets) * 3 +
        min(basic.named_range_count, 30) * 0.4 +
        min(risk.cross_sheet_reference_count, 500) * 0.04
    )
    mc = min(20, int(metadata_score))

    sub = ComplexitySubScores(
        data_scale=ds, formula_depth=fd, vba_mass=vm,
        smell_density=sd, metadata_complexity=mc,
    )
    total = ds + fd + vm + sd + mc

    rationale = {
        "data_scale": f"log10({n_cells}) cells -> {ds}/20",
        "formula_depth": f"max-cond={int(max_cond)}, max-ops={int(max_op)}, max-chain={int(max_chain)} -> {fd}/20",
        "vba_mass": f"{basic.vba_total_lines} LOC across {basic.vba_module_count} modules -> {vm}/20",
        "smell_density": f"{len(smells)} smells / {basic.cell_count_nonempty} cells -> {sd}/20",
        "metadata_complexity": (
            f"{len(risk.hidden_sheets)} hidden + {len(risk.very_hidden_sheets)} veryHidden + "
            f"{basic.named_range_count} named ranges + {risk.cross_sheet_reference_count} "
            f"cross-sheet refs -> {mc}/20"
        ),
    }
    return ComplexityScore(total=total, sub_scores=sub, rationale=dict(sorted(rationale.items())))


# =============================================================================
# Main entry
# =============================================================================

def build_audit(path: Path, sanitize: bool = False) -> AuditReport:
    parse_errors: list = []
    path = Path(path)

    cell_rows, sheet_meta, named_ranges, cf_total = extract_cells(path, parse_errors, sanitize=sanitize)
    vba_modules, vba_text = extract_vba(path, parse_errors)

    n_total = len(cell_rows)
    n_formula = sum(1 for cr in cell_rows if cr.formula)
    n_unique = len({cr.value for cr in cell_rows if cr.value})
    sheet_state_count: Counter = Counter(m["state"] for m in sheet_meta.values())
    dv_total = sum(m["dv_count"] for m in sheet_meta.values())
    vba_total_lines = sum(m.line_count for m in vba_modules)

    basic = BasicStats(
        sheet_count=len(sheet_meta),
        sheet_count_visible=sheet_state_count.get("visible", 0),
        sheet_count_hidden=sheet_state_count.get("hidden", 0),
        sheet_count_very_hidden=sheet_state_count.get("veryHidden", 0),
        cell_count_nonempty=n_total,
        cell_count_formula=n_formula,
        cell_count_unique_values=n_unique,
        named_range_count=len(named_ranges),
        conditional_formatting_count=cf_total,
        data_validation_count=dv_total,
        vba_module_count=len(vba_modules),
        vba_total_lines=vba_total_lines,
        parse_errors_count=len(parse_errors),
    )

    sheets: list = []
    for name in sorted(sheet_meta.keys()):
        m = sheet_meta[name]
        sheets.append(SheetSummary(
            name=name, state=m["state"],
            rows_used=m["rows_used"], cols_used=m["cols_used"],
            cells_nonempty=m["n_nonempty"], cells_formula=m["n_formula"],
            max_ref=m["max_ref"],
            conditional_formatting_count=m["cf_count"],
            data_validation_count=m["dv_count"],
        ))

    sheet_names = set(sheet_meta.keys())
    multi_ref_findings, outgoing, incoming = detect_multiple_references(cell_rows, sheet_names)
    chain_findings = detect_long_calculation_chains(cell_rows, outgoing)
    cond_findings = detect_conditional_complexity(cell_rows)
    multi_op_findings = detect_multiple_operations(cell_rows)
    magic_smell_findings, magic_index = detect_magic_numbers(cell_rows, vba_text)
    dup_findings, pattern_to_cells = detect_duplicated_formulas(cell_rows)

    all_smells = (
        multi_ref_findings + chain_findings + cond_findings +
        multi_op_findings + magic_smell_findings + dup_findings
    )
    smell_type_order = {t: i for i, t in enumerate(SMELL_TYPES)}
    all_smells.sort(key=lambda s: (smell_type_order.get(s.smell_type, 99), -s.metric, s.location))

    # NEW: A.1 / A.2 / A.3
    pillars = detect_pillars(cell_rows, incoming, named_ranges=named_ranges)
    anomalies = detect_anomalies(pattern_to_cells)
    vba_classifications = classify_modules(vba_modules)

    # Build sheet-to-sheet edge weights for the Mermaid sheet-flow diagram.
    # incoming map: target 'sheet|REF' -> set of source 'sheet|ref' formula cells.
    # We aggregate to (src_sheet, tgt_sheet) -> count of distinct source cells.
    sheet_edges: dict = defaultdict(int)
    edge_sources_seen: dict = defaultdict(set)
    for tgt_key, srcs in incoming.items():
        if "|" not in tgt_key:
            continue
        tgt_sheet = tgt_key.split("|", 1)[0]
        for src_key in srcs:
            if "|" not in src_key:
                continue
            src_sheet = src_key.split("|", 1)[0]
            if src_sheet == tgt_sheet:
                continue  # only cross-sheet edges
            # Count each distinct source-formula-cell once per edge
            edge_key = (src_sheet, tgt_sheet)
            if src_key not in edge_sources_seen[edge_key]:
                edge_sources_seen[edge_key].add(src_key)
                sheet_edges[edge_key] += 1
    # Convert to plain dict with stable ordering for determinism.
    sheet_edges_dict = dict(sorted(sheet_edges.items(), key=lambda kv: (kv[0][0], kv[0][1])))

    risk = _build_risk_indicators(cell_rows, sheet_meta, parse_errors, outgoing)
    complexity = _compute_complexity(basic, sheets, all_smells, risk)

    meta = FileMeta(
        file_name=path.name,
        file_size_bytes=path.stat().st_size,
        sha256=_sha256(path),
    )

    methodology = {
        "library_versions": {
            "openpyxl": openpyxl.__version__,
            "oletools": _OLETOOLS_VERSION,
            "formulas": _FORMULAS_VERSION,
        },
        "smell_thresholds": {
            "multiple-references": THRESH_MULTIPLE_REFS,
            "long-calculation-chain": THRESH_LONG_CHAIN,
            "conditional-complexity": THRESH_CONDITIONAL_NESTING,
            "multiple-operations": THRESH_MULTIPLE_OPS,
            "duplicated-formulas": THRESH_DUPLICATED_PATTERN,
        },
        "logic_depth_thresholds": {
            "pillar-fanin-min": 20,
            "pillar-top-n": 10,
            "anomaly-cluster-min-size": 5,
            "anomaly-outlier-fraction": 0.05,
        },
        "confidence_semantics": {
            "high": "exact deterministic count or test (e.g. fan-in count)",
            "medium": "tokenizer-based analysis with well-defined rules",
            "low": "statistical or skipped analysis (e.g. workbook too large)",
        },
        "trivial_numbers_filter": sorted(TRIVIAL_NUMBERS),
        "vba_external_keywords": list(VBA_EXTERNAL_KEYWORDS),
        "vba_classifier_categories": [
            "data-loader", "transformer", "report-writer",
            "ui-handler", "dead-suspected", "mixed",
        ],
        "sanitize_mode": bool(sanitize),
    }

    report = AuditReport(
        meta=meta, basic_stats=basic, sheets=sheets, named_ranges=named_ranges,
        complexity=complexity, smells=all_smells, magic_numbers=magic_index,
        vba_modules=vba_modules,
        vba_classifications=vba_classifications,
        pillars=pillars,
        anomalies=anomalies,
        risk_indicators=risk, methodology=methodology,
        sanitized=bool(sanitize),
        domain_hint=None,  # set below — needs the partly-built report
        workflow=None,
        vba_narratives=None,
        domain_template_matches=None,
        _sheet_edges=sheet_edges_dict,
    )

    # Domain hint runs over the assembled report (it scans sheet/named-range/VBA names)
    report.domain_hint = detect_domain(report)

    # Round-3: workflow detection + VBA narration + domain templates
    workflow = detect_workflow(path, vba_modules)
    report.workflow = workflow
    report.vba_narratives = narrate_modules(
        vba_modules, vba_classifications, workflow_steps=workflow.get("steps") or [],
    )
    report.domain_template_matches = evaluate_domain_templates(report)

    return report
