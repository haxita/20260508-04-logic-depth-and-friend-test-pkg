"""Formula tokenization helpers (shared by smells / pillars / anomalies).

Wraps the `formulas` library's tokenizer and exposes:
- _tokenize / _token_class / _token_attr
- _conditional_nesting_depth, _operator_count, _normalize_pattern
- _extract_ranges, _expand_range_to_cells
- _col_letter
"""

from __future__ import annotations

import re
import warnings

# Silence openpyxl-related warnings reachable from formulas import path
warnings.filterwarnings("ignore", message="Unknown extension is not supported")
warnings.filterwarnings("ignore", message="Workbook contains no default style")

from formulas import Parser as FormulaParser  # noqa: E402

CONDITIONAL_FUNCTIONS = frozenset({
    "IF", "IFS", "IFERROR", "IFNA", "CHOOSE", "SWITCH",
})

_FORMULA_PARSER = FormulaParser()


def _col_letter(n: int) -> str:
    """1-indexed column → A1 letters."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _tokenize(formula: str) -> tuple[list, bool]:
    """
    Returns (tokens, ok). ok=False on any tokenizer error.
    Tokens are formulas-library Token objects.
    """
    if not formula or not formula.startswith("="):
        return [], False
    try:
        tokens, _ast = _FORMULA_PARSER.ast(formula)
        return list(tokens), True
    except Exception:
        return [], False


def _token_class(tok) -> str:
    return type(tok).__name__


def _token_attr(tok) -> dict:
    return getattr(tok, "attr", {}) or {}


def _conditional_nesting_depth(tokens: list) -> int:
    """Walk linear token stream; return max paren-stack depth scoped to conditional functions."""
    if not tokens:
        return 0
    cond_stack = 0
    max_depth = 0
    paren_stack = []
    pending_function_is_cond = False

    for tok in tokens:
        cls = _token_class(tok)
        attr = _token_attr(tok)
        name = attr.get("name", "")

        if cls == "Function":
            pending_function_is_cond = (name in CONDITIONAL_FUNCTIONS)
        elif cls == "Parenthesis":
            if name == "(":
                paren_stack.append(pending_function_is_cond)
                if pending_function_is_cond:
                    cond_stack += 1
                    if cond_stack > max_depth:
                        max_depth = cond_stack
                pending_function_is_cond = False
            elif name == ")":
                if paren_stack:
                    was_cond = paren_stack.pop()
                    if was_cond:
                        cond_stack = max(0, cond_stack - 1)
        else:
            pending_function_is_cond = False

    return max_depth


def _operator_count(tokens: list) -> int:
    n = 0
    for tok in tokens:
        cls = _token_class(tok)
        if cls == "OperatorToken":
            n += 1
        elif cls == "Function":
            n += 1
    return n


def _normalize_pattern(tokens: list) -> str:
    """Replace concrete Range refs with `R`, Numbers with `N`, Strings with `S`."""
    parts = []
    for tok in tokens:
        cls = _token_class(tok)
        attr = _token_attr(tok)
        name = attr.get("name", "")
        if cls == "Range":
            parts.append("R")
        elif cls == "Number":
            parts.append("N")
        elif cls == "String":
            parts.append("S")
        elif cls == "Function":
            parts.append(name.upper())
        elif cls == "OperatorToken":
            parts.append(name)
        elif cls == "Parenthesis":
            parts.append(name)
        elif cls == "Separator":
            parts.append(",")
        else:
            parts.append(cls)
    return "".join(parts)


def _extract_ranges(tokens: list, default_sheet: str, sheet_names: set) -> list:
    """Returns list[{sheet, ref, is_cross_sheet}]."""
    out = []
    for tok in tokens:
        if _token_class(tok) != "Range":
            continue
        attr = _token_attr(tok)
        sheet = attr.get("sheet", "") or attr.get("sheet_id", "") or default_sheet
        ref = attr.get("ref", "") or ""
        if not ref:
            continue
        is_cross = bool(attr.get("sheet")) and (sheet != default_sheet)
        out.append({"sheet": sheet, "ref": ref, "is_cross_sheet": is_cross})
    return out


def _expand_range_to_cells(ref: str, sheet: str, max_expand: int = 200) -> list:
    """Expand 'A1:C3' to individual A1-style refs, capped at max_expand."""
    if ":" not in ref:
        return [ref.upper().replace("$", "")]
    a, b = ref.upper().replace("$", "").split(":", 1)
    if not re.search(r"\d", a) or not re.search(r"\d", b):
        return [f"{a}:{b}"]
    m1 = re.match(r"([A-Z]+)(\d+)", a)
    m2 = re.match(r"([A-Z]+)(\d+)", b)
    if not (m1 and m2):
        return [f"{a}:{b}"]
    c1, r1 = m1.group(1), int(m1.group(2))
    c2, r2 = m2.group(1), int(m2.group(2))

    def col_to_n(s: str) -> int:
        n = 0
        for ch in s:
            n = n * 26 + (ord(ch) - 64)
        return n

    n1 = col_to_n(c1)
    n2 = col_to_n(c2)
    if n1 > n2:
        n1, n2 = n2, n1
    if r1 > r2:
        r1, r2 = r2, r1

    total = (n2 - n1 + 1) * (r2 - r1 + 1)
    if total > max_expand:
        return [f"{a}:{b}"]
    out = []
    for c in range(n1, n2 + 1):
        col_str = _col_letter(c)
        for r in range(r1, r2 + 1):
            out.append(f"{col_str}{r}")
    return out
