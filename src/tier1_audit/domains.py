"""Domain-specific templates (D7) — manufacturing / logistics / inventory.

Sits ON TOP of the existing `domain.py` keyword detector. When the detector
matches a domain with at least medium confidence, this module checks the
already-extracted artifacts (pillar cells, magic-number index, sheet names,
VBA modules) against domain-specific patterns:

    - common_hardcode_risks: typical hardcoded constants the engineer might
      have buried (capacity ceilings, shift hours, lead times, …). We scan
      the magic-number index + pillar cell labels for matches.
    - scheduling_methods_to_check: well-known algorithm hallmarks. We look
      for keyword fingerprints in VBA module names + Sub names.
    - common_sheet_roles: expected sheet name patterns (BOM, MRP, MPS, …).
      We confirm/note any present, and report any expected ones missing.

This is a pure HEURISTIC layer. We do not pretend to understand business logic
— we just cross-check known industry vocabulary against the workbook's
structural facts. Track B (LLM) will replace the heuristics with semantic
inference.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


# Three templates as required by the task spec. Add more later as needed.
DOMAIN_TEMPLATES = {
    "manufacturing-capacity-planning": {
        "match_domain_names": ["capacity-planning", "operations-s&op"],
        "business_friendly_name": "Manufacturing — Capacity Planning",
        "keywords": ["BOM", "MRP", "MPS", "产能", "排产", "工序", "shift",
                     "throughput", "capacity"],
        "common_sheet_roles": {
            "BOM": "bill of materials, typically 2-5 tier tree",
            "MRP": "material requirements after netting against inventory",
            "MPS": "master production schedule by period",
            "Capacity": "available production capacity per period or shift",
            "Inventory": "current stock balances",
        },
        "common_hardcode_risks": [
            ("capacity ceiling per shift", [r"\bcap[a-z_]*\s*=", "产能上限",
                                             r"capacity\s*[:=]"]),
            ("shift hours (typically 7.5 or 8)", [r"\b7\.?5\b", r"\b8\b\s*\*",
                                                   "班次小时", "shift_hours"]),
            ("lead time (in days)", ["lead_time", "提前期", "leadtime"]),
            ("yield loss percentage", ["yield", "良率", "loss_rate", "LOSS_RATE"]),
        ],
        "scheduling_methods_to_check": [
            ("Round-robin (heuristic)", ["round.?robin", "modulo", "sequence"]),
            ("Clarke-Wright (rare in Excel)", ["clarke.?wright", "savings"]),
            ("Solver-based optimization", ["Solver\\.xlam", "SolverOk", "SolverSolve"]),
        ],
    },
    "logistics-routing": {
        "match_domain_names": ["logistics-routing"],
        "business_friendly_name": "Logistics — Vehicle Routing",
        "keywords": ["route", "vehicle", "dispatch", "shipment", "delivery",
                     "配送", "运输", "carrier", "lane", "freight"],
        "common_sheet_roles": {
            "Routes": "route definitions",
            "Vehicles": "fleet master data",
            "Shipments": "load assignment per route",
            "Carriers": "carrier master with rates",
            "Lanes": "origin/destination corridors",
        },
        "common_hardcode_risks": [
            ("vehicle capacity", ["vehicle_capacity", "max_load", "trailer"]),
            ("max stops per route", ["max_stops", "stop_limit"]),
            ("speed assumption (km/h)", ["speed", r"\b40\b", r"\b60\b", r"\b80\b"]),
            ("service time per stop", ["service_time", "停留", "dwell"]),
        ],
        "scheduling_methods_to_check": [
            ("TSP heuristics", ["tsp", "nearest.?neighbor", "two.?opt"]),
            ("VRP with time windows", ["vrptw", "time.?window"]),
            ("Savings algorithm", ["savings", "clarke.?wright"]),
        ],
    },
    "inventory-supply-chain": {
        "match_domain_names": ["inventory-supply-chain"],
        "business_friendly_name": "Inventory — Supply Chain",
        "keywords": ["inventory", "stock", "reorder", "safety stock", "库存",
                     "lead time", "供应商", "supplier", "EOQ", "ABC"],
        "common_sheet_roles": {
            "Inventory": "current stock per SKU",
            "Suppliers": "supplier master with lead times",
            "Reorder": "computed reorder points",
            "ABC": "ABC classification by usage value",
        },
        "common_hardcode_risks": [
            ("safety stock multiplier", ["safety_stock", "safety_factor", "z_score",
                                          "安全库存"]),
            ("service level z-score", ["service_level", r"\b1\.?6[45]\b",
                                        r"\b1\.?9[68]\b", r"\b2\.?32\b"]),
            ("carrying cost rate", ["carrying_cost", "holding_cost"]),
            ("MOQ", ["MOQ", "min_order_qty", "最小订单"]),
        ],
        "scheduling_methods_to_check": [
            ("EOQ formula", ["EOQ", "economic.?order"]),
            ("(s,S) policy", [r"\bs.?S\b", "reorder_point", "order_up_to"]),
            ("Periodic review", ["periodic.?review", "review_period"]),
        ],
    },
}


@dataclass
class DomainTemplateMatch:
    """One template's check results, ready for rendering."""
    template_key: str
    business_friendly_name: str
    matched_keywords: list = field(default_factory=list)
    sheet_role_hits: list = field(default_factory=list)        # [(role, sheet_name)]
    sheet_role_misses: list = field(default_factory=list)       # [role, ...]
    hardcode_risk_hits: list = field(default_factory=list)      # [(risk_label, evidence_str), ...]
    method_hits: list = field(default_factory=list)             # [(method_label, evidence_str), ...]
    confidence: str = "low"


def _scan_for_patterns(text_blob: str, patterns: list) -> list:
    """Return a list of patterns that fired, with a sample matched substring each."""
    hits: list = []
    for pat in patterns:
        try:
            r = re.compile(pat, re.IGNORECASE)
            m = r.search(text_blob)
            if m:
                hits.append((pat, m.group(0)))
        except re.error:
            continue
    return hits


def _build_haystack(report) -> str:
    """Concat strings from the report for case-insensitive scanning."""
    parts: list = []
    for s in report.sheets:
        parts.append(s.name)
    for nr in report.named_ranges:
        parts.append(nr.name)
        parts.append(nr.ref or "")
    for vm in (getattr(report, "vba_modules", None) or []):
        parts.append(vm.name)
        for sf in vm.sub_functions:
            parts.append(sf.name)
        for rl in vm.range_literals:
            parts.append(rl)
    for pillar in (getattr(report, "pillars", None) or []):
        parts.append(pillar.location)
        parts.append(pillar.row_header)
        parts.append(pillar.col_header)
        parts.append(pillar.named_range)
        parts.append(pillar.value)
    for mn in (getattr(report, "magic_numbers", None) or []):
        parts.append(mn.value)
        parts.append(mn.sample_context)
    return "\n".join(p for p in parts if p)


def _match_sheets_for_role(sheets: list, role_hint: str) -> list:
    """Find sheet names that approximately match a role hint (case-insensitive substring)."""
    hint_lower = role_hint.lower()
    out: list = []
    for s in sheets:
        if hint_lower in s.name.lower():
            out.append(s.name)
    return out


def evaluate_domain_templates(report) -> list:
    """Return matched DomainTemplateMatch records.

    Only returns templates whose `match_domain_names` overlap with the detected
    domain hint, OR whose keywords fire ≥ 2 times in the workbook (fallback for
    cases where domain.py picked a different label).
    """
    domain_hint = getattr(report, "domain_hint", None)
    detected_name = getattr(domain_hint, "domain", "unknown") if domain_hint else "unknown"

    haystack = _build_haystack(report)
    matches: list = []

    for tmpl_key, tmpl in DOMAIN_TEMPLATES.items():
        eligible = detected_name in tmpl["match_domain_names"]
        # Fallback: count keyword hits independently — in case domain.py chose
        # a different (related) label.
        kw_hits = []
        for kw in tmpl["keywords"]:
            if re.search(r"\b" + re.escape(kw) + r"\b", haystack, re.IGNORECASE):
                kw_hits.append(kw)
        if not eligible and len(kw_hits) < 2:
            continue

        # Sheet roles
        role_hits: list = []
        role_misses: list = []
        for role, _desc in tmpl["common_sheet_roles"].items():
            found = _match_sheets_for_role(report.sheets, role)
            if found:
                role_hits.append((role, ", ".join(found)))
            else:
                role_misses.append(role)

        # Hardcode risks
        risk_hits: list = []
        for label, patterns in tmpl["common_hardcode_risks"]:
            hits = _scan_for_patterns(haystack, patterns)
            if hits:
                evidence = ", ".join(f"`{h[1]}`" for h in hits[:3])
                risk_hits.append((label, evidence))

        # Scheduling methods
        method_hits: list = []
        for label, patterns in tmpl["scheduling_methods_to_check"]:
            hits = _scan_for_patterns(haystack, patterns)
            if hits:
                evidence = ", ".join(f"`{h[1]}`" for h in hits[:3])
                method_hits.append((label, evidence))

        # Confidence: more hits = higher
        total_evidence = len(kw_hits) + len(role_hits) + len(risk_hits) + len(method_hits)
        if total_evidence >= 8:
            conf = "high"
        elif total_evidence >= 4:
            conf = "medium"
        else:
            conf = "low"

        matches.append(DomainTemplateMatch(
            template_key=tmpl_key,
            business_friendly_name=tmpl["business_friendly_name"],
            matched_keywords=sorted(kw_hits),
            sheet_role_hits=sorted(role_hits, key=lambda x: x[0]),
            sheet_role_misses=sorted(role_misses),
            hardcode_risk_hits=sorted(risk_hits, key=lambda x: x[0]),
            method_hits=sorted(method_hits, key=lambda x: x[0]),
            confidence=conf,
        ))

    matches.sort(key=lambda m: (m.template_key,))
    return matches
