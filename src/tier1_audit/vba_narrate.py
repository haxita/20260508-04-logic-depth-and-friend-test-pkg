"""VBA module heuristic narration (D4) — template-based, NOT semantic.

We honestly cannot understand what a Sub *does* business-wise from static
analysis alone — that's Track B (LLM augmentation). But we CAN narrate the
*structural* facts in human-readable prose, ordered by call dependency:

    "Module_Main (transformer, 180 lines). Reads from sheets X, Y, Z;
     writes to sheet M; calls helpers A, B. Contains 3 nested loops. Has
     an On Error Resume Next at line 47 — silent failure risk."

The narrative is fully deterministic and does not pretend to know meaning.
We add `<!-- LLM-AUGMENT: vba-narration:<module> -->` markers so a future
Track B step can replace the template prose with a richer LLM summary while
keeping the structural skeleton (and not requiring report restructure).

Order: BFS from button-bound + event-handler entry points (the workflow
"reachable set"); modules unreachable from any entry are flagged as
**possibly dead code**.
"""

from __future__ import annotations

import re
from collections import defaultdict, deque
from dataclasses import dataclass, field

# Re-use the already-tested patterns from vba_classify
_RE_FOR_LOOP = re.compile(r"(?im)^\s*For\b")
_RE_DO_LOOP = re.compile(r"(?im)^\s*Do\b")
_RE_WHILE_LOOP = re.compile(r"(?im)^\s*While\b")
_RE_OER = re.compile(r"(?im)\bOn\s+Error\s+Resume\s+Next\b")
_RE_OER_LINE = re.compile(r"(?i)\bOn\s+Error\s+Resume\s+Next\b")


@dataclass
class VbaNarrative:
    module_name: str
    inferred_type: str          # from VbaClassification
    confidence: str
    line_count: int
    sub_count: int
    func_count: int
    reads_sheets: list = field(default_factory=list)
    writes_sheets: list = field(default_factory=list)
    external_calls: bool = False
    has_oer: bool = False
    oer_lines: list = field(default_factory=list)  # line numbers
    nested_loops_max: int = 0
    callers: list = field(default_factory=list)    # other module names that invoke this one
    callees: list = field(default_factory=list)    # other module names this one invokes
    reachable_from_entry: bool = False
    role_inference: str = ""    # 1-line role hint: 'system entry point', 'helper utility', etc
    notable_patterns: list = field(default_factory=list)  # bullet strings for the report
    narrative: str = ""         # final prose paragraph


def _count_nested_loops(code: str) -> int:
    """Approximate the maximum loop-nesting depth in this module."""
    depth = 0
    max_depth = 0
    for line in code.splitlines():
        s = line.strip()
        if _RE_FOR_LOOP.match(s) or _RE_DO_LOOP.match(s) or _RE_WHILE_LOOP.match(s):
            depth += 1
            max_depth = max(max_depth, depth)
        elif s.lower().startswith(("next", "loop", "wend")):
            depth = max(0, depth - 1)
    return max_depth


def _oer_line_numbers(code: str, max_to_report: int = 3) -> list:
    """Return up to N (1-based) line numbers where 'On Error Resume Next' appears."""
    out: list = []
    for i, line in enumerate(code.splitlines(), start=1):
        if _RE_OER_LINE.search(line):
            out.append(i)
            if len(out) >= max_to_report:
                break
    return out


def _build_inter_module_call_graph(modules: list, sub_to_module: dict) -> tuple:
    """Build (callers_of[mod], callees_of[mod]) maps.

    A module M calls module N when M's source contains a `Call SubXyz` or
    `XYZ()` invocation, and SubXyz is defined in N.
    """
    callers: dict = defaultdict(set)
    callees: dict = defaultdict(set)
    sub_pat = re.compile(r"\b([A-Za-z_]\w*)\s*\(")
    call_kw = re.compile(r"(?im)\bCall\s+([A-Za-z_]\w*)")
    for m in modules:
        code = getattr(m, "source_text", "") or ""
        # Strip body of `Call XYZ` and bare `XYZ(...)` invocations
        candidates: set = set()
        for cm in call_kw.finditer(code):
            candidates.add(cm.group(1))
        for cm in sub_pat.finditer(code):
            candidates.add(cm.group(1))
        for sub in candidates:
            target_mod = sub_to_module.get(sub)
            if target_mod and target_mod != m.name:
                callees[m.name].add(target_mod)
                callers[target_mod].add(m.name)
    return callers, callees


def _build_sub_to_module_map(modules: list) -> dict:
    """Map each defined Sub/Function name to its owning module."""
    out: dict = {}
    sub_def = re.compile(
        r"(?im)^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
        r"(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
        r"([A-Za-z_][\w]*)"
    )
    for m in modules:
        code = getattr(m, "source_text", "") or ""
        for match in sub_def.finditer(code):
            name = match.group(2)
            # First-wins to keep determinism
            if name not in out:
                out[name] = m.name
    return out


def _bfs_reachable(start_set: set, callees: dict) -> set:
    """BFS through the inter-module call graph from `start_set`."""
    seen: set = set()
    q = deque(start_set)
    while q:
        n = q.popleft()
        if n in seen:
            continue
        seen.add(n)
        for nxt in callees.get(n, set()):
            if nxt not in seen:
                q.append(nxt)
    return seen


def _infer_role(narr: VbaNarrative, modules_count: int) -> str:
    """One-line role hint based on structural features."""
    n = narr
    if not n.reachable_from_entry and not n.callers:
        return "isolated module — not reached from any button or event handler (possibly dead)"
    if n.callers and not n.callees:
        return "leaf helper — invoked by other modules but invokes nothing further"
    if not n.callers and n.callees:
        return "system entry point — invokes other modules and is not invoked back"
    if n.line_count > 500 and n.sub_count + n.func_count > 5:
        return "large multi-purpose module — likely the workbook's main logic block"
    if n.inferred_type == "data-loader":
        return "data ingest — reads external sources or workbook ranges into VBA structures"
    if n.inferred_type == "transformer":
        return "compute step — reads inputs, derives outputs, writes back to sheets"
    if n.inferred_type == "report-writer":
        return "output writer — populates sheet ranges with computed values"
    if n.inferred_type == "ui-handler":
        return "UI event handler — fires when the user opens the workbook or interacts with a sheet"
    if n.inferred_type == "dead-suspected":
        return "near-empty shell — typical for default class/sheet stubs with no real code"
    return "mixed responsibilities — no single structural signal dominates"


def _build_notable_patterns(narr: VbaNarrative, vm) -> list:
    """Extract bullet strings about quirks worth flagging."""
    out: list = []
    if narr.has_oer:
        if narr.oer_lines:
            line_str = ", ".join(str(n) for n in narr.oer_lines)
            out.append(
                f"Contains `On Error Resume Next` at line(s) {line_str} — "
                f"silent failure risk; errors are suppressed silently."
            )
        else:
            out.append("Contains `On Error Resume Next` — silent failure risk.")
    if vm.external_keywords:
        kws = ", ".join(f"`{k}`" for k in vm.external_keywords[:3])
        more = f" (+{len(vm.external_keywords)-3} more)" if len(vm.external_keywords) > 3 else ""
        out.append(f"Uses external/COM API keyword(s): {kws}{more}.")
    if narr.nested_loops_max >= 3:
        out.append(
            f"Contains {narr.nested_loops_max} levels of nested loops — "
            f"typically indicates a row × column × period scan pattern."
        )
    return out


def narrate_modules(modules: list, classifications: list,
                    workflow_steps: list = None) -> list:
    """Build VbaNarrative records, ordered by call dependency.

    Args:
      modules: list[VbaModule]
      classifications: list[VbaClassification]
      workflow_steps: list[WorkflowStep] from workflow.detect_workflow()

    Returns:
      Ordered list of VbaNarrative. Order is: BFS from entry-point modules
      (modules hosting user-callable Subs), then alphabetical for unreachable
      modules. Each unreachable module is flagged in narrative.
    """
    cls_by_name = {c.module_name: c for c in classifications}
    sub_to_module = _build_sub_to_module_map(modules)
    callers, callees = _build_inter_module_call_graph(modules, sub_to_module)

    # Entry modules: those owning a Sub used by a workflow step (button or event)
    entry_modules: set = set()
    if workflow_steps:
        for step in workflow_steps:
            if step.module_name:
                entry_modules.add(step.module_name)
            elif step.sub_name and step.sub_name in sub_to_module:
                entry_modules.add(sub_to_module[step.sub_name])

    reachable = _bfs_reachable(entry_modules, callees) if entry_modules else set()

    narratives: list = []
    for vm in modules:
        cls = cls_by_name.get(vm.name)
        code = getattr(vm, "source_text", "") or ""
        n_sub = sum(1 for sf in vm.sub_functions if sf.kind == "Sub")
        n_func = sum(1 for sf in vm.sub_functions if sf.kind == "Function")
        narr = VbaNarrative(
            module_name=vm.name,
            inferred_type=cls.inferred_type if cls else "mixed",
            confidence=cls.confidence if cls else "low",
            line_count=vm.line_count,
            sub_count=n_sub,
            func_count=n_func,
            reads_sheets=list(cls.reads_sheets if cls else []),
            writes_sheets=list(cls.writes_sheets if cls else []),
            external_calls=bool(cls.external_calls if cls else False),
            has_oer=bool(vm.has_on_error_resume_next),
            oer_lines=_oer_line_numbers(code),
            nested_loops_max=_count_nested_loops(code),
            callers=sorted(callers.get(vm.name, set())),
            callees=sorted(callees.get(vm.name, set())),
            reachable_from_entry=(vm.name in reachable or vm.name in entry_modules),
        )
        narr.role_inference = _infer_role(narr, len(modules))
        narr.notable_patterns = _build_notable_patterns(narr, vm)
        narratives.append(narr)

    # Build prose paragraphs
    for narr in narratives:
        narr.narrative = _render_narrative_prose(narr)

    # Order: entry modules first (BFS depth), then unreachable, then alphabetical
    depth: dict = {}
    visited: set = set()
    q: deque = deque((m, 0) for m in sorted(entry_modules))
    while q:
        m, d = q.popleft()
        if m in visited:
            continue
        visited.add(m)
        depth[m] = d
        for c in callees.get(m, ()):
            if c not in visited:
                q.append((c, d + 1))

    def sort_key(narr: VbaNarrative):
        if narr.module_name in depth:
            return (0, depth[narr.module_name], narr.module_name.lower())
        if narr.reachable_from_entry:
            return (1, 0, narr.module_name.lower())
        return (2, 0, narr.module_name.lower())

    narratives.sort(key=sort_key)
    return narratives


def _render_narrative_prose(n: VbaNarrative) -> str:
    """Build the multi-line prose paragraph for one module."""
    lines: list = []
    # Header line
    role = n.role_inference
    lines.append(f"**Role inference**: {role}.")

    # What it does (structural)
    parts: list = []
    if n.reads_sheets:
        sheets = ", ".join(f"`{s}`" for s in n.reads_sheets[:5])
        more = f" (+{len(n.reads_sheets)-5} more)" if len(n.reads_sheets) > 5 else ""
        parts.append(f"reads from sheets {sheets}{more}")
    if n.writes_sheets:
        sheets = ", ".join(f"`{s}`" for s in n.writes_sheets[:5])
        more = f" (+{len(n.writes_sheets)-5} more)" if len(n.writes_sheets) > 5 else ""
        parts.append(f"writes to sheets {sheets}{more}")
    if n.callees:
        c = ", ".join(f"`{c}`" for c in n.callees[:3])
        more = f" (+{len(n.callees)-3} more)" if len(n.callees) > 3 else ""
        parts.append(f"calls into modules {c}{more}")
    if n.callers and not n.callees:
        c = ", ".join(f"`{c}`" for c in n.callers[:3])
        parts.append(f"is invoked by modules {c}")
    if n.nested_loops_max >= 2:
        parts.append(f"contains {n.nested_loops_max} levels of nested loops")
    if not parts:
        parts.append("contains no salient sheet I/O or inter-module calls")
    lines.append("**What it does (structural)**: " + "; ".join(parts) + ".")

    # Notable patterns
    if n.notable_patterns:
        lines.append("**Notable patterns**:")
        for bp in n.notable_patterns:
            lines.append(f"- {bp}")

    # Call relationships
    rel_parts: list = []
    if n.callers:
        c = ", ".join(f"`{c}`" for c in n.callers[:3])
        more = f" (+{len(n.callers)-3} more)" if len(n.callers) > 3 else ""
        rel_parts.append(f"called by {c}{more}")
    else:
        rel_parts.append("not called by any other module")
    if n.callees:
        c = ", ".join(f"`{c}`" for c in n.callees[:3])
        more = f" (+{len(n.callees)-3} more)" if len(n.callees) > 3 else ""
        rel_parts.append(f"calls {c}{more}")
    lines.append("**Call relationships**: " + "; ".join(rel_parts) + ".")

    if not n.reachable_from_entry and n.module_name not in {"ThisWorkbook.cls"}:
        lines.append("**Possibly dead code** — no path reaches this module from any "
                     "button or event handler we detected.")

    return "\n".join(lines)
