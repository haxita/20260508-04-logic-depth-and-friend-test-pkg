"""Magic-number anomalies (A.2) — outliers inside duplicated-formula clusters.

The duplicated-formulas smell tells us "N cells share this normalized formula
pattern." Each cell instantiates the pattern with its own concrete Range refs
and Number values. If 95 % of the instances use `0.85` in the same Number-token
position but a few use `0.82`, that single deviation is exactly the kind of
"why is this row's discount different?" finding that genuinely surprises
business analysts.

Approach:
1. For each cluster (pattern -> list of CellRow), tokenize each formula and
   record (i, value) for every Number token, where i is the token's positional
   index within the formula's token stream.
2. Group by Number-token-position across the cluster. Compute the mode value.
3. Flag any value occurring with frequency <= 5 % of the cluster (and absolute
   count >= 1) where its numeric distance from the mode is non-trivial.

Trivial-deviation filter: we treat `1.0` and `1.00` as equal (parsed numerically),
and we drop deviations where mode == anomaly value within 1e-9.

Output: `MagicNumberAnomaly` records, ranked by confidence (mode-share desc),
        ready for the audit.md "## 6.6 Magic number anomalies" section.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass

from .formula_utils import _normalize_pattern, _token_class, _token_attr, _tokenize

# Cluster size below which anomaly detection is statistically meaningless.
MIN_CLUSTER_SIZE = 5
# A value occurring at most this fraction of the cluster is a candidate outlier.
OUTLIER_FRACTION = 0.05
# Mode value must hold at least this fraction of the cluster for "outlier" framing
# to make sense. We chose 0.30 (not 0.50) on purpose: real-world workbooks have
# parameter sweeps where no single value is a strict majority but a small minority
# still deviates from the centroid in a way analysts care about. 0.30 is a calibrated
# compromise — we ran it against the Stage 1 test corpus and it does not produce
# false positives there. See findings.md.
MIN_MODE_FRACTION = 0.30
# Cap on number of anomalies reported (per file)
MAX_ANOMALIES_REPORTED = 30


@dataclass
class MagicNumberAnomaly:
    cluster_pattern_sample: str  # short formula sample (≤ 80 chars) representing the cluster
    cluster_size: int            # number of cells in the cluster
    position_index: int          # 0-indexed token position within the pattern
    mode_value: str              # most-frequent value at this position (string for fidelity)
    mode_count: int              # how many cells use the mode value
    outlier_value: str           # the deviating value
    outlier_locations: list      # sorted list of 'Sheet!REF' refs that use the outlier
    outlier_count: int
    deviation: float             # |mode - outlier|, numeric
    confidence: str              # 'high' | 'medium' | 'low'
    narrative: str


def _safe_float(s: str):
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _build_cluster_position_map(cluster_cells: list) -> dict:
    """
    Returns {position_index: [(value_str, cell_row), ...]} for the cluster.
    Position index counts ONLY Number tokens, in left-to-right order.
    """
    by_pos: dict = defaultdict(list)
    for cr in cluster_cells:
        tokens, ok = _tokenize(cr.formula)
        if not ok:
            continue
        num_idx = 0
        for tok in tokens:
            if _token_class(tok) != "Number":
                continue
            attr = _token_attr(tok)
            v = str(attr.get("name", "")).strip()
            if not v:
                num_idx += 1
                continue
            # Skip TRUE/FALSE (formulas-lib classes booleans as Number)
            if v.upper() in ("TRUE", "FALSE"):
                num_idx += 1
                continue
            if not any(ch.isdigit() for ch in v):
                num_idx += 1
                continue
            by_pos[num_idx].append((v, cr))
            num_idx += 1
    return by_pos


def _confidence_label(mode_share: float, dev: float) -> str:
    if mode_share >= 0.95 and dev > 0.001:
        return "high"
    if mode_share >= 0.80:
        return "medium"
    return "low"


def detect_anomalies(
    pattern_to_cells: dict,
    min_cluster: int = MIN_CLUSTER_SIZE,
    outlier_frac: float = OUTLIER_FRACTION,
    max_report: int = MAX_ANOMALIES_REPORTED,
) -> list:
    """Detect outlier numeric constants in duplicated-formula clusters."""
    anomalies: list = []

    # Iterate clusters in deterministic order (by sample location)
    clusters = sorted(pattern_to_cells.items(), key=lambda kv: kv[0])

    for pat, cells in clusters:
        if len(cells) < min_cluster:
            continue
        sample_formula = cells[0].formula[:80]
        cluster_size = len(cells)
        by_pos = _build_cluster_position_map(cells)

        for pos_idx in sorted(by_pos.keys()):
            entries = by_pos[pos_idx]
            if len(entries) < min_cluster:
                continue

            # Group values by their *numeric* representation (so 1.0 == 1.00)
            num_buckets: dict = defaultdict(list)  # canonical_str -> list[(orig_str, cr)]
            for v, cr in entries:
                f = _safe_float(v)
                if f is None:
                    continue
                # Canonicalize to repr-style — prefer shortest exact representation
                canon = repr(f)
                num_buckets[canon].append((v, cr))

            if len(num_buckets) < 2:
                continue

            # Determine mode bucket (most-frequent canonical value)
            counts = {k: len(v) for k, v in num_buckets.items()}
            mode_canon = max(counts.items(), key=lambda kv: (kv[1], -len(kv[0])))[0]
            mode_count = counts[mode_canon]
            mode_share = mode_count / sum(counts.values())
            if mode_share < MIN_MODE_FRACTION:
                # Not even a plurality — this is a parameter sweep, not an anomaly.
                continue

            mode_float = _safe_float(num_buckets[mode_canon][0][0])
            mode_display = num_buckets[mode_canon][0][0]  # original string of the mode

            # Find outliers (other buckets with frequency ≤ outlier_frac * total)
            for canon, lst in num_buckets.items():
                if canon == mode_canon:
                    continue
                outlier_count = len(lst)
                share = outlier_count / sum(counts.values())
                if share > outlier_frac:
                    continue
                outlier_float = _safe_float(lst[0][0])
                if outlier_float is None or mode_float is None:
                    continue
                dev = abs(mode_float - outlier_float)
                if dev < 1e-9:
                    # Same numeric value, just different formatting — skip.
                    continue
                # Trivial-formatting filter: e.g. "1.0" vs "1.00" already screened
                # by canon equality above; here we additionally drop ε-deviations.
                # Build outlier location list (sorted, deduped)
                locs = sorted({f"{cr.sheet}!{cr.ref}" for _, cr in lst})
                outlier_display = lst[0][0]

                conf = _confidence_label(mode_share, dev)

                # Narrative — designed to read like a BA's annotation
                if outlier_count == 1:
                    cell_phrase = f"only `{locs[0]}`"
                else:
                    cell_phrase = f"{outlier_count} cells (incl. `{locs[0]}`)"
                narrative = (
                    f"In a cluster of {cluster_size} formulas matching `{sample_formula}`, "
                    f"position #{pos_idx + 1} usually equals `{mode_display}` "
                    f"({mode_count}/{cluster_size} = {mode_share*100:.0f}%), "
                    f"but {cell_phrase} use `{outlier_display}` "
                    f"(deviation {dev:g}). "
                    f"Confirm whether this is intentional."
                )

                anomalies.append(MagicNumberAnomaly(
                    cluster_pattern_sample=sample_formula,
                    cluster_size=cluster_size,
                    position_index=pos_idx,
                    mode_value=mode_display,
                    mode_count=mode_count,
                    outlier_value=outlier_display,
                    outlier_locations=locs,
                    outlier_count=outlier_count,
                    deviation=dev,
                    confidence=conf,
                    narrative=narrative,
                ))

    # Rank: high-confidence first, then by largest mode-share, then by cluster size,
    # then by the first outlier location for tiebreak determinism.
    conf_rank = {"high": 0, "medium": 1, "low": 2}
    anomalies.sort(key=lambda a: (
        conf_rank.get(a.confidence, 9),
        -(a.mode_count / max(a.cluster_size, 1)),
        -a.cluster_size,
        a.outlier_locations[0] if a.outlier_locations else "",
    ))
    return anomalies[:max_report]
