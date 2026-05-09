"""Lightweight domain hint detector — pure keyword matching, zero LLM.

Scans sheet names, named ranges, VBA Sub/Function names, and the most
frequent identifier-like tokens in formulas. For each predefined domain
it counts how many keywords from that domain's lexicon are matched, then
returns the top match plus a confidence band.

Decisions:
- Word-boundary matching (case-insensitive). Substring matches would over-fire.
- Multiple matches of the same keyword in different sources count once each
  (we count *how many distinct keywords* from the domain were seen, not the
  raw frequency). This biases toward genuine domain coverage rather than a
  single repetition.
- "Unknown" domain when no keyword fires anywhere — we don't fabricate.

Display in the executive summary in render.py / render_html.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass


# Each value is a list of keywords. Add liberally; word-boundary keeps it sane.
# Plural forms are listed alongside singular because real-world sheets
# typically use plural ("Routes", "Vehicles") and `\b...\b` does not match a
# substring across word characters.
DOMAINS = {
    "capacity-planning": [
        "BOM", "MRP", "MPS", "产能", "排产", "工序", "shift", "throughput",
        "capacity", "工时", "节拍",
    ],
    "logistics-routing": [
        "route", "routes", "vehicle", "vehicles", "dispatch", "shipment",
        "shipments", "delivery", "deliveries", "fleet", "carrier", "carriers",
        "lane", "lanes", "freight",
        "配送", "运输", "TMS", "车辆", "路径", "线路",
    ],
    "inventory-supply-chain": [
        "inventory", "stock", "reorder", "safety stock", "EOQ",
        "movements", "movement", "supplier", "suppliers",
        "库存", "lead time", "供应商", "采购", "安全库存", "在库",
    ],
    "operations-s&op": [
        "forecast", "demand", "sales", "revenue", "season", "S&OP",
        "需求", "预测",
    ],
    "actuarial-insurance": [
        "actuarial", "reserve", "premium", "claim",
        "保单", "保费", "annuity", "policy", "理赔",
    ],
    "financial-modeling": [
        "NPV", "IRR", "DCF", "discount", "valuation", "WACC",
        "cash flow", "现金流", "估值",
    ],
}

# Keep the iteration order deterministic for narrative stability.
DOMAIN_ORDER = list(DOMAINS.keys())


@dataclass
class DomainHint:
    domain: str            # e.g. 'capacity-planning' or 'unknown'
    confidence: str        # 'high' | 'medium' | 'low' | 'none'
    hits: int              # how many distinct keywords matched
    matched_keywords: list  # the actual keywords found, sorted lex
    runner_up: str         # 2nd-best domain name or '' if none
    runner_up_hits: int    # 2nd-best count
    rationale: str         # human-readable one-liner


def _word_pattern(keyword: str) -> re.Pattern:
    """Build a case-insensitive word-boundary regex for a keyword.

    For ASCII keywords we use \\b on each side. For CJK / mixed-script
    keywords, Python's \\b doesn't cleanly mark boundaries between CJK
    and ASCII, so we use lookarounds against word-character classes
    that include CJK ranges. The simpler choice: re.escape the keyword
    and let \\b apply on the ASCII side; for pure-CJK keywords, \\b is
    effectively a no-op and substring matching takes over (acceptable
    for short distinctive Chinese terms like 产能 / 排产).
    """
    esc = re.escape(keyword)
    # If the keyword starts/ends with an ASCII word char, anchor with \b.
    # Otherwise (pure-CJK) skip anchoring since \b doesn't help.
    left = r"\b" if keyword[0].isascii() and keyword[0].isalnum() else ""
    right = r"\b" if keyword[-1].isascii() and keyword[-1].isalnum() else ""
    return re.compile(left + esc + right, re.IGNORECASE)


# Pre-compile all patterns once
_COMPILED: dict = {
    domain: [(kw, _word_pattern(kw)) for kw in keywords]
    for domain, keywords in DOMAINS.items()
}


_RE_VBA_SUBFUNC_NAME = re.compile(
    r"(?im)^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"([A-Za-z_][\w]*)"
)


def _collect_text_corpus(report) -> str:
    """Concatenate every text source we want to scan into one big haystack.

    Sources:
      - sheet names
      - named-range names
      - VBA Sub/Function names (extracted by regex, since the report has
        them already on each VbaModule.sub_functions list)
      - top-N identifier-like substrings inside formulas (sheet names appear
        in formulas as cross-sheet refs, which often carry domain signal)
    """
    parts: list = []

    for s in report.sheets:
        parts.append(s.name)

    for nr in report.named_ranges:
        parts.append(nr.name)
        parts.append(nr.ref)  # named-range refs include sheet names like _constants!$C$2

    for vm in getattr(report, "vba_modules", []) or []:
        parts.append(vm.name)
        for sf in vm.sub_functions:
            parts.append(sf.name)
        # Range literals often contain user-named sheets / table names
        for rl in vm.range_literals:
            parts.append(rl)

    return "\n".join(parts)


def detect_domain(report) -> DomainHint:
    """Return the best-matching domain with a confidence band."""
    haystack = _collect_text_corpus(report)

    # Per-domain unique keyword matches
    matches: dict = defaultdict(set)
    for domain in DOMAIN_ORDER:
        for kw, pat in _COMPILED[domain]:
            if pat.search(haystack):
                matches[domain].add(kw)

    # Rank by hit count, with deterministic tie-break by DOMAIN_ORDER index
    ranked = sorted(
        (
            (domain, len(matches[domain]), sorted(matches[domain]))
            for domain in DOMAIN_ORDER
        ),
        key=lambda x: (-x[1], DOMAIN_ORDER.index(x[0])),
    )

    top_name, top_hits, top_kws = ranked[0]
    runner_name, runner_hits, _ = ranked[1] if len(ranked) > 1 else ("", 0, [])

    if top_hits == 0:
        return DomainHint(
            domain="unknown",
            confidence="none",
            hits=0,
            matched_keywords=[],
            runner_up="",
            runner_up_hits=0,
            rationale="domain not auto-detected — analyze manually",
        )

    if top_hits >= 5:
        confidence = "high"
    elif top_hits >= 2:
        confidence = "medium"
    else:
        confidence = "low"

    rationale = (
        f"matched {top_hits} keyword(s) from `{top_name}` lexicon: "
        + ", ".join(top_kws)
    )

    return DomainHint(
        domain=top_name,
        confidence=confidence,
        hits=top_hits,
        matched_keywords=top_kws,
        runner_up=runner_name if runner_hits > 0 else "",
        runner_up_hits=runner_hits,
        rationale=rationale,
    )
