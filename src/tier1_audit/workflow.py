"""Workflow detection — buttons, event handlers, VBA call graph (D3).

Pure stdlib: zipfile + xml.etree.ElementTree only. NO new pip deps.

Goal: derive "how the workbook is operationally used" from static analysis.

Inputs available at audit time:
    - The xlsm zip itself (re-opened here to read xl/drawings/* and xl/worksheets/_rels/*)
    - Already-extracted VbaModule list with `source_text` (raw VBA per module)
    - Sheet metadata (for sheet name resolution)

Outputs:
    - WorkflowStep records, ordered by dependency (sub A writes Sheet X, sub B
      reads Sheet X → A before B; ties broken by VBA call dependency).
    - Each step has: button label, bound Sub name, sheets read, sheets written,
      called Subs, source-line-of-evidence ('VBA' or 'sheet rels').
    - When NO buttons + NO event handlers → returns empty list (caller renders
      the "no buttons detected" graceful message).

Static-analysis honesty:
    The xl/drawings parsing is intentionally conservative. We extract:
      - VML drawings (legacy form controls): <x:FmlaMacro> + nearby button text
      - DrawingML <controlPr> macro=... elements (modern form controls embedded
        in worksheets)
      - Sheet-level <control name="..."> nodes that link to controlPr macros
    We do NOT attempt to parse ActiveX OLE objects (that's a binary format).
    If the workbook uses ActiveX controls, those buttons will not be detected
    and the report says so honestly.
"""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict, deque
from dataclasses import dataclass, field
from pathlib import Path

# XML namespaces we need
_NS = {
    "x": "urn:schemas-microsoft-com:office:excel",
    "v": "urn:schemas-microsoft-com:vml",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rels": "http://schemas.openxmlformats.org/package/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}


@dataclass
class WorkflowButton:
    """A user-clickable surface that invokes a VBA Sub."""
    label: str          # button text (or '' if none)
    macro: str          # macro reference, e.g. 'Trend' or 'MapsSheet.FindDrivingDistance'
    sheet: str          # sheet name where the button lives ('' if unknown)
    source: str         # 'vml' | 'controlPr' | 'event-handler'


@dataclass
class WorkflowStep:
    """One operational step the user performs (button click or auto-fired event)."""
    order: int                  # 1-based, after dependency sort
    label: str                  # human-visible button text or event description
    macro: str                  # full macro target (Module.Sub or Sub)
    sub_name: str               # bare Sub name (last segment)
    module_name: str            # owning VBA module ('' if unknown)
    sheet: str                  # sheet the user clicks on
    reads_sheets: list = field(default_factory=list)
    writes_sheets: list = field(default_factory=list)
    calls: list = field(default_factory=list)   # Subs this step calls (for narrative)
    source_kind: str = ""       # 'button' | 'workbook-event' | 'sheet-event'


_RE_FMLA_MACRO = re.compile(r"<x:FmlaMacro[^>]*>([^<]*)</x:FmlaMacro>")
_RE_VML_TEXT = re.compile(r">([^<>]+?)</font>")
_RE_VML_SHAPE = re.compile(r"<v:shape\b", re.IGNORECASE)


def _strip_macro(macro: str) -> tuple:
    """Parse `[0]!ModuleX.SubY` into (module, sub).

    Forms seen in the wild:
        [0]!ModuleX.SubY   — typical
        ModuleX.SubY       — already cleaned
        SubY               — bare
        '[0]!Trend'        — quoted edges
    """
    m = (macro or "").strip().strip("'").strip('"')
    if m.startswith("[0]!"):
        m = m[4:]
    if "!" in m:
        m = m.split("!", 1)[1]
    if "." in m:
        mod, sub = m.split(".", 1)
        return mod, sub
    return "", m


def _parse_vml_buttons(vml_bytes: bytes) -> list:
    """Extract (label, macro_string) pairs from a vmlDrawing*.vml file.

    Returns a list of (label, macro_full_text) tuples (one per button shape).
    """
    text = vml_bytes.decode("utf-8", errors="replace")
    out: list = []
    # Each <v:shape> is a button; FmlaMacro is inside <x:ClientData>.
    # We split on <v:shape and process each chunk.
    parts = re.split(r"(?=<v:shape\b)", text, flags=re.IGNORECASE)
    for part in parts:
        if "<v:shape" not in part.lower():
            continue
        macros = _RE_FMLA_MACRO.findall(part)
        if not macros:
            continue
        # Pick the first non-empty macro
        macro = next((m for m in macros if m.strip()), "")
        if not macro:
            continue
        # Find the button text: <font>...</font> in the textbox
        texts = _RE_VML_TEXT.findall(part)
        label = ""
        if texts:
            label = " ".join(t.strip() for t in texts).strip()
            label = re.sub(r"\s+", " ", label)
        out.append((label, macro))
    return out


def _parse_sheet_controls(sheet_xml: bytes) -> list:
    """Find <control name="..." macro="..."> elements in a sheet.xml.

    These are modern form controls. Returns list of (name, macro) tuples.
    """
    text = sheet_xml.decode("utf-8", errors="replace")
    out: list = []
    # The sheet's control element has a name attr and inside controlPr is a macro attr.
    # We search the text for `<control ` blocks and pull both attributes.
    for m in re.finditer(r"<control\b([^>]*?)>", text):
        attrs = m.group(1)
        name_m = re.search(r'name="([^"]*)"', attrs)
        name = name_m.group(1) if name_m else ""
        # macro is in the controlPr child, search forward
        after = text[m.end(): m.end() + 600]
        macro_m = re.search(r'macro="([^"]*)"', after)
        macro = macro_m.group(1) if macro_m else ""
        if macro:
            out.append((name, macro))
    return out


def _build_sheet_id_map(zf: zipfile.ZipFile) -> dict:
    """Map the sheet's relationship-id chain so we can attach a button to its sheet name.

    Returns {sheetN_xml_path: sheet_name}, e.g. {'xl/worksheets/sheet13.xml': 'Forecast P1 - Trend'}
    """
    mapping: dict = {}
    try:
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8", errors="replace")
        wb_rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="replace")
    except KeyError:
        return mapping
    # rId -> target path (e.g. 'worksheets/sheet1.xml')
    # URLs in attrs contain '/', so we can't use [^/]*? — use a permissive
    # but bounded greedy match with the closing /> as the anchor.
    rid_to_target: dict = {}
    for m in re.finditer(r'<Relationship\b([^>]*?)/?>', wb_rels):
        attrs = m.group(1)
        id_m = re.search(r'\bId="([^"]+)"', attrs)
        target_m = re.search(r'\bTarget="([^"]+)"', attrs)
        if id_m and target_m:
            rid_to_target[id_m.group(1)] = target_m.group(1)
    # sheet name -> rId
    for m in re.finditer(r'<sheet\b([^/]*?)/>', wb_xml):
        attrs = m.group(1)
        name_m = re.search(r'name="([^"]*)"', attrs)
        rid_m = re.search(r'r:id="([^"]+)"', attrs)
        if not name_m or not rid_m:
            continue
        sheet_name = name_m.group(1)
        rid = rid_m.group(1)
        target = rid_to_target.get(rid, "")
        if target:
            full_path = "xl/" + target if not target.startswith("xl/") else target
            mapping[full_path] = sheet_name
    return mapping


def _resolve_sheet_for_drawing(zf: zipfile.ZipFile, drawing_target: str,
                               sheet_id_map: dict) -> str:
    """Find the sheet name that references a given drawing/vmlDrawing path.

    `drawing_target` is the path under xl/, e.g. 'xl/drawings/vmlDrawing2.vml'.
    """
    target_basename = drawing_target.replace("xl/drawings/", "")
    for sheet_path in sheet_id_map.keys():
        rels_path = sheet_path.replace(".xml", ".xml.rels").replace(
            "xl/worksheets/", "xl/worksheets/_rels/"
        )
        try:
            rels = zf.read(rels_path).decode("utf-8", errors="replace")
        except KeyError:
            continue
        if target_basename in rels:
            return sheet_id_map.get(sheet_path, "")
    return ""


def detect_buttons(xlsm_path) -> list:
    """Return list[WorkflowButton]. Empty if nothing detected."""
    buttons: list = []
    path = Path(xlsm_path)
    if not path.exists():
        return buttons
    try:
        zf = zipfile.ZipFile(str(path), "r")
    except (zipfile.BadZipFile, OSError):
        return buttons
    try:
        sheet_id_map = _build_sheet_id_map(zf)
        # 1. Look for sheet-level <control macro="..."> elements (modern controls).
        #    These are tied to a specific sheet.
        for sheet_path, sheet_name in sheet_id_map.items():
            try:
                sheet_xml = zf.read(sheet_path)
            except KeyError:
                continue
            controls = _parse_sheet_controls(sheet_xml)
            for ctrl_name, macro in controls:
                buttons.append(WorkflowButton(
                    label=ctrl_name or "(unnamed control)",
                    macro=macro,
                    sheet=sheet_name,
                    source="controlPr",
                ))
        # 2. Walk vmlDrawings — only count those LINKED via a sheet's rels.
        # (Orphan VML drawings — present in the zip but not referenced — are
        # historical artifacts and should not be reported as live buttons.)
        linked_vml: set = set()
        for sheet_path in sheet_id_map.keys():
            rels_path = sheet_path.replace(".xml", ".xml.rels").replace(
                "xl/worksheets/", "xl/worksheets/_rels/"
            )
            try:
                rels = zf.read(rels_path).decode("utf-8", errors="replace")
            except KeyError:
                continue
            for m in re.finditer(r'Target="([^"]*?vmlDrawing[^"]*?)"', rels):
                vml_target = m.group(1)
                # Resolve relative path
                if vml_target.startswith("../"):
                    full = "xl/" + vml_target[3:]
                else:
                    full = vml_target
                linked_vml.add(full)

        for vml_path in sorted(linked_vml):
            try:
                vml_bytes = zf.read(vml_path)
            except KeyError:
                continue
            sheet_name = _resolve_sheet_for_drawing(zf, vml_path, sheet_id_map)
            for label, macro in _parse_vml_buttons(vml_bytes):
                # If a controlPr button already exists with the same sheet+macro,
                # upgrade its label using the VML one (VML text is human-friendly).
                upgraded = False
                for b in buttons:
                    if b.sheet == sheet_name and _strip_macro(b.macro)[1] == _strip_macro(macro)[1]:
                        if label and (not b.label or b.label.lower().startswith("button")):
                            b.label = label
                        upgraded = True
                        break
                if upgraded:
                    continue
                buttons.append(WorkflowButton(
                    label=label or "(button)",
                    macro=macro,
                    sheet=sheet_name,
                    source="vml",
                ))
    finally:
        try:
            zf.close()
        except Exception:
            pass
    # Stable sort: by sheet, then label
    buttons.sort(key=lambda b: (b.sheet.lower(), b.label.lower(), b.macro))
    return buttons


# =============================================================================
# VBA call graph + event handler detection
# =============================================================================

# Sub/Function definition (with module name will be attached externally)
_RE_SUB_DEF = re.compile(
    r"(?im)^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
    r"(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"([A-Za-z_][\w]*)"
)
# `Call XYZ` and `XYZ(...)`-style invocations
_RE_CALL_KEYWORD = re.compile(r"(?im)\bCall\s+([A-Za-z_][\w\.]*)")
_RE_BARE_INVOKE = re.compile(r"(?im)^\s*([A-Za-z_][\w\.]*)\s*\(?\s*[A-Za-z_\"\d]")
# Event handlers
_RE_EVENT_HANDLER = re.compile(
    r"(?i)^(Workbook_(Open|BeforeClose|BeforeSave|SheetChange|NewSheet)|"
    r"Worksheet_(Change|SelectionChange|Activate|BeforeDoubleClick)|"
    r"\w+_(Click|Change))$"
)
# Sheet reads/writes inside VBA
_RE_SHEET_REF = re.compile(r'(?i)\b(?:Sheets|Worksheets)\s*\(\s*"([^"]+)"\s*\)')
_RE_RANGE_VALUE_WRITE = re.compile(
    r"""(?ix)
    \b(?:Range|Cells)\s*\([^)]*\)\s*(?:\.\s*Value)?\s*=
    """
)
# Implicit host-sheet write when no Sheets() qualifier appears: matches
#   Range(...) = ...   |   Cells(...) = ...   |   Me.[Foo] = ...
_RE_IMPLICIT_WRITE = re.compile(
    r"""(?ix)
    (?:^|\n|\:|\s)
    (?:Range\s*\([^)]*\)|Cells\s*\([^)]*\)|Me\s*\.\s*\[[^\]]+\]|Me\s*\.\s*[A-Za-z_]\w*)
    \s*(?:\.\s*Value)?\s*=
    """
)
_RE_IMPLICIT_READ = re.compile(
    r"""(?ix)
    =\s*[^=].*?
    (?:Range\s*\([^)]*\)|Cells\s*\([^)]*\)|Me\s*\.\s*\[[^\]]+\]|Me\s*\.\s*[A-Za-z_]\w*)
    """
)


def _module_subs(modules: list) -> dict:
    """Build a lookup of sub_name -> (module_name, source_text).

    When two modules define the same sub name, the one with longer text is
    chosen (heuristic: backup__/legacy modules vs the actual one).
    """
    lookup: dict = {}
    for m in modules:
        code = getattr(m, "source_text", "") or ""
        for match in _RE_SUB_DEF.finditer(code):
            sub_name = match.group(2)
            if sub_name in lookup:
                # Prefer the larger module (most likely the canonical one)
                if len(code) <= len(lookup[sub_name][2]):
                    continue
            lookup[sub_name] = (m.name, sub_name, code)
    return lookup


def _classify_event(sub_name: str) -> str:
    """Return 'workbook-event' / 'sheet-event' / '' for a given sub name."""
    s = sub_name.lower()
    if s.startswith("workbook_"):
        return "workbook-event"
    if s.startswith("worksheet_"):
        return "sheet-event"
    if "_click" in s or "_change" in s:
        return "sheet-event"
    return ""


def _vba_extract_step_facts(sub_name: str, code: str, default_sheet: str = "") -> dict:
    """Extract (reads_sheets, writes_sheets, calls) from a Sub's body.

    Heuristics:
      - Explicit `Sheets("X")` references near a Value-write count as `X` writes.
      - Implicit `Range(...)` or `Me.[...]` writes (no qualifier) count as writes
        to `default_sheet` (the host sheet of the button or .cls module).
      - All other Sheets() references count as reads.
    """
    # Locate the Sub body
    sub_pat = re.compile(rf"(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?"
                         rf"(?:Static\s+)?(?:Sub|Function)\s+{re.escape(sub_name)}\b"
                         rf"(.+?)(?=^\s*End\s+(?:Sub|Function)\b)", re.DOTALL)
    m = sub_pat.search(code)
    body = m.group(1) if m else code  # fallback to whole module

    writes_sheets: set = set()
    reads_sheets: set = set()
    for wm in _RE_RANGE_VALUE_WRITE.finditer(body):
        ctx = body[max(0, wm.start() - 200): wm.end() + 50]
        sheet_refs_in_ctx = list(_RE_SHEET_REF.finditer(ctx))
        if sheet_refs_in_ctx:
            for sm in sheet_refs_in_ctx:
                writes_sheets.add(sm.group(1))
        elif default_sheet:
            writes_sheets.add(default_sheet)

    # Pure implicit writes (Me.[X] = ..., Range(...) = ... without explicit Sheets())
    if default_sheet and _RE_IMPLICIT_WRITE.search(body):
        writes_sheets.add(default_sheet)
    if default_sheet and _RE_IMPLICIT_READ.search(body):
        reads_sheets.add(default_sheet)

    all_refs = {sm.group(1) for sm in _RE_SHEET_REF.finditer(body)}
    reads_sheets |= (all_refs - writes_sheets)

    calls: set = set()
    for cm in _RE_CALL_KEYWORD.finditer(body):
        calls.add(cm.group(1).split(".")[-1])
    # Bare invocations are noisy. We pick up Module.Sub() patterns conservatively:
    # match `\b\w+\.\w+\s*\(` as a callee reference.
    for cm in re.finditer(r"\b([A-Za-z_]\w*)\.([A-Za-z_]\w*)\s*\(", body):
        # filter out common false positives
        receiver = cm.group(1)
        method = cm.group(2)
        if receiver in {"Range", "Cells", "Sheets", "Worksheets", "Me", "Application",
                         "Workbook", "Worksheet", "ActiveSheet", "ActiveWorkbook",
                         "Selection", "Target", "Response", "Err", "Debug", "MsgBox"}:
            continue
        if method in {"Value", "Text", "Cells", "Range", "Item", "Add", "Count",
                       "Offset", "Resize", "End", "Find", "ToString", "Trim"}:
            continue
        calls.add(method)

    return {
        "reads": sorted(reads_sheets),
        "writes": sorted(writes_sheets),
        "calls": sorted(calls),
    }


def detect_event_handlers(modules: list) -> list:
    """Return a list of WorkflowButton-like records for event handlers (no button).

    Event handlers are auto-fired when the user opens the workbook or edits a
    sheet — they're operationally relevant even without a button click.
    """
    out: list = []
    for m in modules:
        code = getattr(m, "source_text", "") or ""
        for match in _RE_SUB_DEF.finditer(code):
            kind = match.group(1).lower()
            if kind != "sub":
                continue
            sub_name = match.group(2)
            ev = _classify_event(sub_name)
            if not ev:
                continue
            label = sub_name
            out.append(WorkflowButton(
                label=label, macro=f"{m.name}.{sub_name}",
                sheet="", source=f"event-handler:{ev}",
            ))
    out.sort(key=lambda b: (b.macro.lower(),))
    return out


def build_workflow_steps(buttons: list, event_handlers: list, modules: list) -> list:
    """Topologically sort entry points by sheet-write/read dependency + VBA calls.

    Entry points = buttons + event handlers. For each, we resolve:
        - sub body
        - reads_sheets / writes_sheets
        - calls (other Subs by name)

    Then we toposort:
        - If step A writes Sheet X and step B reads Sheet X → A before B.
        - If A calls something that B's sub eventually does → A and B are on the
          same chain (we keep the one with deeper write count first).
        - Else, lexicographic by macro for determinism.
    """
    sub_lookup = _module_subs(modules)
    entries: list = []
    for entry in buttons + event_handlers:
        mod_part, sub_part = _strip_macro(entry.macro)
        # Resolve to actual module if missing
        info = sub_lookup.get(sub_part)
        if info:
            module_name, _, source_text = info
        elif mod_part:
            # Try a module that ends with `.cls` / `.bas` matching mod_part
            match_mod = next(
                (m for m in modules if m.name.lower().startswith(mod_part.lower())),
                None,
            )
            if match_mod is None:
                continue
            module_name = match_mod.name
            source_text = getattr(match_mod, "source_text", "") or ""
        else:
            continue
        facts = _vba_extract_step_facts(sub_part, source_text, default_sheet=entry.sheet)
        # Source kind
        if entry.source.startswith("event-handler"):
            sk = entry.source.split(":", 1)[1] if ":" in entry.source else "event-handler"
        else:
            sk = "button"
        step = WorkflowStep(
            order=0,  # filled below
            label=entry.label or sub_part,
            macro=entry.macro,
            sub_name=sub_part,
            module_name=module_name,
            sheet=entry.sheet,
            reads_sheets=facts["reads"],
            writes_sheets=facts["writes"],
            calls=facts["calls"],
            source_kind=sk,
        )
        entries.append(step)

    # Build dependency edges: A -> B (A must run before B) iff A writes a sheet B reads
    edges: dict = defaultdict(set)  # node_idx -> set of dependents
    indeg: dict = defaultdict(int)
    nodes = list(range(len(entries)))
    for i, a in enumerate(entries):
        for j, b in enumerate(entries):
            if i == j:
                continue
            if set(a.writes_sheets) & set(b.reads_sheets):
                if j not in edges[i]:
                    edges[i].add(j)
                    indeg[j] += 1
    # Initialize indeg for all nodes
    for n in nodes:
        if n not in indeg:
            indeg[n] = 0

    # Kahn's algorithm with deterministic tie-break (macro lex)
    ready = sorted([n for n in nodes if indeg[n] == 0],
                   key=lambda n: entries[n].macro.lower())
    order: list = []
    seen: set = set()
    while ready:
        n = ready.pop(0)
        if n in seen:
            continue
        seen.add(n)
        order.append(n)
        for nxt in sorted(edges[n], key=lambda x: entries[x].macro.lower()):
            indeg[nxt] -= 1
            if indeg[nxt] == 0:
                ready.append(nxt)
        ready.sort(key=lambda x: entries[x].macro.lower())
    # Append leftovers (cyclic / disconnected)
    for n in nodes:
        if n not in seen:
            order.append(n)

    out: list = []
    for i, idx in enumerate(order, start=1):
        s = entries[idx]
        s.order = i
        out.append(s)
    return out


def detect_workflow(xlsm_path, vba_modules: list) -> dict:
    """End-to-end: detect buttons + event handlers + topologically sort steps.

    Returns:
        {
            "buttons": [WorkflowButton, ...],
            "event_handlers": [WorkflowButton, ...],
            "steps": [WorkflowStep, ...],
            "no_buttons_detected": bool,
        }
    """
    buttons = detect_buttons(xlsm_path)
    handlers = detect_event_handlers(vba_modules)
    steps = build_workflow_steps(buttons, handlers, vba_modules)
    return {
        "buttons": buttons,
        "event_handlers": handlers,
        "steps": steps,
        "no_buttons_detected": (len(buttons) == 0 and len(handlers) == 0),
    }
