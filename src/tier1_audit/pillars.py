"""Pillar cells (A.1) — Top-N fan-in narratives, with column-block dedupe.

A "pillar cell" is one that many formulas depend on. Modifying it has
high blast-radius; it is one of the most useful diagnostic facts a BA
can have when triaging a workbook nobody else can maintain.

Polish-round upgrade: when many cells in the same column share the same
fan-in and same affected-sheet set (a typical "data column referenced
symmetrically by N MPS formulas" pattern), we collapse them into a
single group entry with `member_count` instead of bloating Top-20 with
identical-pattern rows. After dedupe we surface Top-10 *distinct* groups
(falling back to single-cell entries when no grouping applies).

Input: the `incoming` map already built by smells.detect_multiple_references
       (target 'sheet|REF' -> set of source 'sheet|ref' formula cells).
       Plus the cell_rows so we can tell whether the pillar itself is a
       formula or a value (constant), and which sheets the dependents live in.

Output: `Pillar` records, sorted by descending fan-in. Each record has a
        plain-language narrative line so the audit.md report can include
        it verbatim. A `member_count` of 1 means an individual cell; >1
        means a deduped group of equivalent column cells.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass, field

# fan-in threshold for "pillar" — distinct from the multiple-references smell threshold.
# We want pillars to be a small, high-impact list, not a long-tail of medium hubs.
PILLAR_FANIN_THRESHOLD = 20
# After dedupe we surface fewer entries; the original 20 was inflated by clones.
PILLAR_TOP_N = 10
# Initial scan size before dedupe — collect more, then collapse, then truncate.
PILLAR_SCAN_TOP_N = 100


@dataclass
class Pillar:
    location: str            # representative 'Sheet!REF' (or 'Sheet!Cstart..Cend' for groups)
    fan_in: int              # number of distinct dependent formula cells (per member)
    affected_sheets: list    # sorted list of sheet names that contain dependents
    affected_sheet_count: int
    is_formula_itself: bool  # True if the pillar cell is itself a formula
    sample_dependents: list  # up to 5 sample 'Sheet!REF' refs of dependents (sorted)
    narrative: str           # human-readable single-sentence explanation
    pillar_kind: str         # 'constant-input' | 'formula-relay' | 'whole-column' | 'column-block'
    member_count: int = 1    # 1 = individual cell; >1 = collapsed column-block group
    member_refs: list = field(default_factory=list)  # all cell refs in this group, sorted
    # Round-3 inline value+label fields (Michael's #1 critique fix):
    # the report must show what _constants!C4 actually IS, not just the ref.
    value: str = ""          # the cell's actual value (or first cell's, for column-block)
    value_kind: str = ""     # 'number' | 'text' | 'formula' | 'empty'
    row_header: str = ""     # leftmost text-bearing label in same row (heuristic)
    col_header: str = ""     # topmost text-bearing label in same column (heuristic)
    named_range: str = ""    # named range that resolves to this cell (or '')


_RE_CELL_REF = re.compile(r"^([A-Z]+)(\d+)$")


def _parse_cell_ref(ref: str):
    """Return (col_letter, row_int) or None if not a single A1 ref."""
    m = _RE_CELL_REF.match(ref.upper().replace("$", ""))
    if not m:
        return None
    return m.group(1), int(m.group(2))


def _pillar_kind(target_key: str, formula_cells: set) -> str:
    if "|" in target_key and ":" in target_key.split("|", 1)[1]:
        return "whole-column"
    return "formula-relay" if target_key in formula_cells else "constant-input"


def _truncate_label(s: str, max_len: int = 40) -> str:
    """Trim long labels for inline display while keeping the gist."""
    if not s:
        return ""
    s = str(s).replace("\n", " ").replace("\r", " ").strip()
    if len(s) <= max_len:
        return s
    return s[: max_len - 1].rstrip() + "…"


def _value_label_phrase(p: Pillar) -> str:
    """Produce the inline 'value=X, label=Y, named range=Z' clause used in narratives.

    Empty when nothing useful is known (graceful degradation).
    """
    parts: list = []
    if p.value:
        parts.append(f"value `{_truncate_label(p.value, 30)}`")
    if p.named_range:
        parts.append(f"named range `{p.named_range}`")
    label = p.row_header or p.col_header
    if label and label != p.value:
        parts.append(f"label `{_truncate_label(label, 30)}`")
    if not parts:
        return ""
    return " (" + ", ".join(parts) + ")"


def _make_narrative(p: Pillar) -> str:
    """Render the human-language line. Designed to read like a BA's note."""
    n_sheets = len(p.affected_sheets)
    if p.pillar_kind == "constant-input":
        kind_word = "constant input"
    elif p.pillar_kind == "formula-relay":
        kind_word = "formula relay"
    elif p.pillar_kind == "whole-column":
        kind_word = "whole-column reference"
    else:  # column-block
        kind_word = "column block"

    if n_sheets == 1:
        scope_phrase = f"sheet `{p.affected_sheets[0]}`"
    elif n_sheets <= 3:
        scope_phrase = "sheets " + ", ".join(f"`{s}`" for s in p.affected_sheets)
    else:
        listed = ", ".join(f"`{s}`" for s in p.affected_sheets[:3])
        scope_phrase = f"{n_sheets} sheets (incl. {listed})"

    inline = _value_label_phrase(p)

    if p.member_count > 1:
        # Group narrative
        risk = (
            "very-high cumulative blast radius" if p.fan_in >= 100 else
            "high cumulative blast radius" if p.fan_in >= 50 else
            "notable cumulative blast radius"
        )
        return (
            f"{p.member_count} cells in this column block{inline} each cascade into "
            f"{p.fan_in} formulas across {scope_phrase} — modifying the column header "
            f"or any single cell propagates similarly ({risk})."
        )

    # Individual cell narrative (legacy)
    risk = (
        "very-high-risk change point" if p.fan_in >= 100 else
        "high-risk change point" if p.fan_in >= 50 else
        "notable change point"
    )
    return (
        f"Modifying `{p.location}`{inline} (a {kind_word}) would cascade into "
        f"{p.fan_in} formulas across {scope_phrase} — {risk}."
    )


def _summarize_member_range(refs_sorted: list) -> str:
    """Given a sorted list of 'Col<row>' refs, build a concise range string.

    Single ref:  'C42'
    All same column, contiguous rows:  'C10..C115 (106 cells)'
    Otherwise:  'C10, C12, C14 (and 8 more)'
    """
    if not refs_sorted:
        return ""
    if len(refs_sorted) == 1:
        return refs_sorted[0]

    parsed = [_parse_cell_ref(r) for r in refs_sorted]
    if all(p is not None for p in parsed):
        cols = {p[0] for p in parsed}
        if len(cols) == 1:
            col = parsed[0][0]
            rows = sorted(p[1] for p in parsed)
            lo, hi = rows[0], rows[-1]
            # If most rows in [lo, hi] are present, show as a column block
            # range; otherwise list a few + 'and N more'.
            density = len(rows) / max(hi - lo + 1, 1)
            if density >= 0.5:
                return f"{col}{lo}..{col}{hi} ({len(rows)} cells)"
    # Fallback: list first 3 + count
    if len(refs_sorted) <= 3:
        return ", ".join(refs_sorted)
    return f"{', '.join(refs_sorted[:3])} (+{len(refs_sorted) - 3} more)"


def _group_key(target_key: str, fan_in: int, affected_sheets: list) -> tuple:
    """Build the dedupe key for a candidate cell.

    Two candidates are equivalent (collapse-able) when they:
      - Live on the same sheet
      - Are in the same column (column letter)
      - Have identical fan_in
      - Have identical affected_sheets set

    Whole-column refs (Cstart:Cend) and non-A1 references don't dedupe.
    """
    if "|" not in target_key:
        return ("ungroupable", target_key)
    sheet, ref = target_key.split("|", 1)
    parsed = _parse_cell_ref(ref)
    if parsed is None:
        # Whole-column or other multi-cell ref — never dedupe
        return ("ungroupable", target_key)
    col, _row = parsed
    return ("group", sheet, col, fan_in, frozenset(affected_sheets))


def _build_cell_lookups(cell_rows: list) -> tuple:
    """Build the lookups we need to resolve pillar value + headers.

    Returns (cells_by_key, sheet_row_text, sheet_col_text):
      - cells_by_key: 'sheet|REF' -> CellRow
      - sheet_row_text: (sheet, row) -> [(col, text)] sorted by col asc
      - sheet_col_text: (sheet, col) -> [(row, text)] sorted by row asc
    """
    cells_by_key: dict = {}
    sheet_row_text: dict = defaultdict(list)
    sheet_col_text: dict = defaultdict(list)
    for cr in cell_rows:
        cells_by_key[f"{cr.sheet}|{cr.ref}"] = cr
        # Heuristic: only non-formula text-bearing cells count as labels.
        # Formulas can produce text too, but we prefer raw labels (typical Excel pattern).
        if cr.formula:
            continue
        if not cr.value:
            continue
        # Skip pure-number-like values when used as label seed
        v = cr.value.strip()
        if not v:
            continue
        # Numeric cell? value parses cleanly as float? — still allowed (e.g. row 1
        # often has a year), but de-prioritized when picking the leftmost label
        sheet_row_text[(cr.sheet, cr.row)].append((cr.col, v))
        sheet_col_text[(cr.sheet, cr.col)].append((cr.row, v))
    for k in sheet_row_text:
        sheet_row_text[k].sort(key=lambda x: x[0])
    for k in sheet_col_text:
        sheet_col_text[k].sort(key=lambda x: x[0])
    return cells_by_key, sheet_row_text, sheet_col_text


_RE_NAMED_RANGE_REF = re.compile(r"^\s*'?([^'!]+?)'?!?\$?([A-Z]+)\$?(\d+)\s*$", re.IGNORECASE)
_RE_NAMED_RANGE_COLREF = re.compile(r"^\s*'?([^'!]+?)'?!\$?([A-Z]+):\$?([A-Z]+)\s*$", re.IGNORECASE)


def _build_named_range_lookup(named_ranges: list) -> dict:
    """Map 'sheet|REF' -> name for cell-level resolution.

    Whole-column refs like '_constants!$C:$C' map every cell in that column;
    we record the column-only key 'sheet|COL' separately so the pillar resolver
    can fall back when the exact ref doesn't match.
    """
    by_cell: dict = {}
    by_col: dict = {}
    for nr in named_ranges or []:
        ref = (getattr(nr, "ref", "") or "").strip()
        if not ref:
            continue
        # Sometimes destinations is multi: split on comma.
        for part in ref.split(","):
            part = part.strip()
            if not part:
                continue
            m = _RE_NAMED_RANGE_REF.match(part)
            if m:
                sheet, col, row = m.group(1), m.group(2).upper(), int(m.group(3))
                by_cell.setdefault(f"{sheet}|{col}{row}", nr.name)
                continue
            m2 = _RE_NAMED_RANGE_COLREF.match(part)
            if m2 and m2.group(2).upper() == m2.group(3).upper():
                # whole-column named range
                sheet, col = m2.group(1), m2.group(2).upper()
                by_col.setdefault(f"{sheet}|{col}", nr.name)
    return {"by_cell": by_cell, "by_col": by_col}


def _resolve_value_and_labels(
    sheet: str, ref: str,
    cells_by_key: dict,
    sheet_row_text: dict,
    sheet_col_text: dict,
    nr_lookup: dict,
) -> dict:
    """Look up the pillar's value, row header, col header, named range.

    Heuristics:
      - row_header: leftmost text-bearing label in the same row at col 1 or 2
                    (preferring text over numbers); fallback to any leftmost text
                    on that row to the LEFT of the pillar cell.
      - col_header: topmost text-bearing label in the same column at row 1 or 2
                    (preferring text); fallback to any topmost text ABOVE pillar.
    """
    parsed = _parse_cell_ref(ref)
    if parsed is None:
        # whole-column or other non-A1 ref
        return {"value": "", "value_kind": "", "row_header": "",
                "col_header": "", "named_range": ""}
    col_letter, row_num = parsed
    # Translate col_letter -> column index (1-based)
    col_idx = 0
    for ch in col_letter:
        col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)

    # Value
    cr = cells_by_key.get(f"{sheet}|{ref}")
    value = ""
    value_kind = "empty"
    if cr is not None:
        if cr.formula:
            # Show the formula but keep it short
            value = cr.formula if len(cr.formula) <= 40 else cr.formula[:37] + "..."
            value_kind = "formula"
        elif cr.value:
            value = cr.value.strip()
            try:
                float(value)
                value_kind = "number"
            except (ValueError, TypeError):
                value_kind = "text"

    # Row header (leftmost text in same row, prefer columns 1-2; non-numeric preferred)
    row_header = ""
    row_cells = sheet_row_text.get((sheet, row_num), [])
    # Prefer columns A/B (1-2), then any leftmost cell strictly to the LEFT of pillar
    for c, text in row_cells:
        if c >= col_idx:
            break
        # prefer text over number; but accept any non-empty
        try:
            float(text)
            num_candidate = text  # remember as fallback
            continue
        except (ValueError, TypeError):
            row_header = text
            break
    else:
        # No text found before pillar; try numeric leftmost
        for c, text in row_cells:
            if c >= col_idx:
                break
            row_header = text
            break

    # Col header (topmost text in same column, prefer rows 1-2; non-numeric preferred)
    col_header = ""
    col_cells = sheet_col_text.get((sheet, col_idx), [])
    for r, text in col_cells:
        if r >= row_num:
            break
        try:
            float(text)
            continue  # skip numeric topmost; keep looking
        except (ValueError, TypeError):
            col_header = text
            break
    else:
        for r, text in col_cells:
            if r >= row_num:
                break
            col_header = text
            break

    # Named range
    nr = nr_lookup["by_cell"].get(f"{sheet}|{ref}", "")
    if not nr:
        nr = nr_lookup["by_col"].get(f"{sheet}|{col_letter}", "")

    return {
        "value": value,
        "value_kind": value_kind,
        "row_header": row_header,
        "col_header": col_header,
        "named_range": nr,
    }


def detect_pillars(
    cell_rows: list, incoming: dict,
    named_ranges: list = None,
    threshold: int = PILLAR_FANIN_THRESHOLD,
    top_n: int = PILLAR_TOP_N,
) -> list:
    """Rank cells by fan-in. Apply column-block dedupe. Return at most `top_n` Pillar records.

    `named_ranges` is the audit's NamedRange list (may be None). When given, each
    pillar attempts to resolve its named range, value, and row/col label.
    """
    # Build a quick lookup: which targets are themselves formula cells?
    formula_cells: set = {f"{cr.sheet}|{cr.ref}" for cr in cell_rows if cr.formula}
    cells_by_key, sheet_row_text, sheet_col_text = _build_cell_lookups(cell_rows)
    nr_lookup = _build_named_range_lookup(named_ranges or [])

    # Filter incoming map by threshold
    candidates: list = []
    for target_key, srcs in incoming.items():
        n = len(srcs)
        if n < threshold:
            continue
        candidates.append((target_key, n, srcs))

    # Sort by (-n, target_key) for stable, deterministic ranking
    candidates.sort(key=lambda x: (-x[1], x[0]))
    # Initial pre-dedupe cap — generous so we don't lose info to top-N filter
    # before we've even collapsed clones.
    candidates = candidates[:PILLAR_SCAN_TOP_N]

    # Group by (sheet, col, fan_in, affected_sheets_frozenset)
    groups: dict = defaultdict(list)
    group_order: list = []  # insertion order = original ranking order
    for target_key, n, srcs in candidates:
        sheet, ref = target_key.split("|", 1)
        src_sheets = sorted({s.split("|", 1)[0] for s in srcs})
        gkey = _group_key(target_key, n, src_sheets)
        if gkey not in groups:
            group_order.append(gkey)
        groups[gkey].append((target_key, n, srcs, sheet, ref, src_sheets))

    # Build Pillar records, one per group
    pillars: list = []
    for gkey in group_order:
        members = groups[gkey]
        # Pick the representative member (first by ranking order = highest-fanin
        # tie-break by lex; for a group they all share fan_in, so this is just
        # the lex-first cell ref).
        rep_target_key, n, srcs, sheet, rep_ref, src_sheets = members[0]

        # Sample dependents come from the representative; for a group the
        # dependents share the same shape so any member's sample is valid.
        sample_dep_keys = sorted(srcs)[:5]
        sample_dependents = [
            f"{k.split('|', 1)[0]}!{k.split('|', 1)[1]}"
            for k in sample_dep_keys
        ]

        # All member refs, sorted by row number (numeric) inside the column,
        # else lex.
        all_refs = []
        for tk, _n, _srcs, _sh, ref, _sh_list in members:
            all_refs.append(ref)
        # Sort: parse row when possible
        def _ref_sort_key(r: str):
            p = _parse_cell_ref(r)
            return (p[0], p[1]) if p else (r, 0)
        all_refs_sorted = sorted(all_refs, key=_ref_sort_key)

        if len(members) > 1 and gkey[0] == "group":
            # Build a column-block representative location
            range_summary = _summarize_member_range(all_refs_sorted)
            location = f"{sheet}!{range_summary}"
            kind = "column-block"
            # For column-blocks we resolve value+labels from the FIRST member ref
            # (its row label often differs per row but value comes from a single cell)
            resolved = _resolve_value_and_labels(
                sheet, all_refs_sorted[0],
                cells_by_key, sheet_row_text, sheet_col_text, nr_lookup,
            )
        else:
            location = f"{sheet}!{rep_ref}"
            kind = _pillar_kind(rep_target_key, formula_cells)
            resolved = _resolve_value_and_labels(
                sheet, rep_ref,
                cells_by_key, sheet_row_text, sheet_col_text, nr_lookup,
            )

        is_formula = (kind == "formula-relay")

        p = Pillar(
            location=location,
            fan_in=n,
            affected_sheets=src_sheets,
            affected_sheet_count=len(src_sheets),
            is_formula_itself=is_formula,
            sample_dependents=sample_dependents,
            narrative="",
            pillar_kind=kind,
            member_count=len(members),
            member_refs=[f"{sheet}!{r}" for r in all_refs_sorted],
            value=resolved["value"],
            value_kind=resolved["value_kind"],
            row_header=resolved["row_header"],
            col_header=resolved["col_header"],
            named_range=resolved["named_range"],
        )
        p.narrative = _make_narrative(p)
        pillars.append(p)

    # After dedupe we re-rank: by descending fan-in, then by descending member_count
    # (a column-block of fan-in 72 with 100 members is more interesting than a
    # single-cell pillar of fan-in 72), then lex by location.
    pillars.sort(key=lambda p: (-p.fan_in, -p.member_count, p.location))
    return pillars[:top_n]
