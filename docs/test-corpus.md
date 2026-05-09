# Test Corpus — Track A regression baseline

Round 4 (Stage 2 expansion). **Audience: contributors, friends running the friend
test, anyone evaluating consistency of the Track A zero-LLM audit pipeline.**

This document captures the canonical inputs, expected outputs, and
cross-fixture invariants for the audit pipeline as of `audit v0.1.0`.
Every number below is a regression target — if a re-run produces different
numbers, either the audit logic changed (intentionally — bump the range) or
the fixture changed (un-intended — investigate before tweaking the test).

---

## 1. Fixture inventory

Six total fixtures, three brought in by Round 4:

| Fixture | Style | Size | Domain (detected) | VBA | Buttons | Notes |
|---|---|---|---|---|---|---|
| `capacity_planning_synth.xlsm` | formula-heavy + VBA donor | 528 KB | `capacity-planning` | 40 mods / 14.7K LOC | 0 | Existing baseline. Manufacturing capacity planning, Chinese-localized sheets (BOM/MPS/MRP/产能规划). |
| `OperationPlanning.xlsm` | real-world spreadsheet | 440 KB | `capacity-planning` | 17 mods / 210 LOC | 8 | Real customer-style file; 358 named ranges; many buttons trigger VBA forecast routines. |
| `vba-web-example.xlsm` | VBA-heavy, low data | 467 KB | `unknown` | 40 mods / 14.7K LOC | 6 | The donor file. Unrelated VBA (HTTP / Twitter API). Confirms `unknown` falls through gracefully. |
| `tests/fixtures/logistics_routing_synth.xlsm` (new) | formula + named ranges + 1 button | 427 KB | `logistics-routing` | 40 mods / 14.7K LOC | 1 | Routes/Vehicles/Shipments/Customers + 订单 (mixed CN/EN). Real `<controlPr macro=>` button on Routes. |
| `tests/fixtures/inventory_supply_synth.xlsm` (new) | formula-rich, EOQ math | 429 KB | `inventory-supply-chain` | 40 mods / 14.7K LOC | 0 | EOQ + safety-stock + reorder formulas. Data validation on Movements. No buttons → graceful "no buttons" path exercised. |
| `tests/fixtures/minimal_no_vba.xlsm` (new) | pure values + 8 formulas | 6 KB | `unknown` | 0 mods | 0 | Degradation/robustness test. No VBA, no named ranges, no hidden sheets, no CF, no DV. |

**On the donor VBA**: the three VBA-bearing synthetic fixtures all reuse the
**same** `vba-web-example.xlsm` donor blob (40 modules, 14.7K LOC of HTTP /
OAuth / Twitter API code). This is documented honestly — the VBA is
*unrelated* to the workbook's data; its only purpose is to populate
`wb.vba_archive` so the audit's classifier has substance to chew on. This
mimics a common real-world pattern: legacy authors copy-paste VBA from
elsewhere and let dead code accumulate.

**Why fixture sizes are ~430 KB**: the donor's `xl/vbaProject.bin` is
~330 KB and dominates. The synthetic data + structure adds the remaining
~100 KB. The `inventory_supply_synth` size landed slightly above the
spec's 200-300 KB target band for that variant — accepted because the
spec explicitly allowed reusing the donor blob ("**Or** use vba-web-example
donor with NO button bindings").

---

## 2. Per-fixture expected audit findings

Numbers are point estimates from the `audit v0.1.0` baseline run. Each
score has a tolerance band that the regression tests in
`tests/test_new_corpus.py` and `tests/test_corpus.py` enforce.

### 2.1 `capacity_planning_synth.xlsm` (existing — unchanged)

| Metric | Value |
|---|---|
| Complexity score | 87 / 100 (band ±5: 82–92) |
| Detected domain | `capacity-planning` (medium, 4 hits) |
| Domain templates matched | `manufacturing-capacity-planning` (high), `inventory-supply-chain` (low — see note) |
| Pillar count (deduped) | 2 |
| Anomaly count | 0 |
| Smell count | 154 (multi-ref 50, magic-numbers 50, multi-op 30, dup 15, cond-complex 9) |
| Workflow Guide buttons | 0 (graceful "formula-driven only") |
| VBA modules | 40 (data-loader 18, dead-suspected 8, mixed 6, report-writer 5, transformer 3) |

**Note on inventory-supply-chain "low" match**: Round 4 added `供应商` /
`库存` keywords for plural matching. Capacity-planning workbooks
legitimately reference suppliers + inventory, so a "low" cross-template
hit is honest and expected. The primary template stays
`manufacturing-capacity-planning` at high confidence.

### 2.2 `OperationPlanning.xlsm` (existing — unchanged)

| Metric | Value |
|---|---|
| Complexity score | 90 / 100 |
| Detected domain | `capacity-planning` (medium, 2 hits) |
| Domain templates matched | `manufacturing-capacity-planning` (medium) |
| Pillar count (deduped) | 10 |
| Anomaly count | 0 |
| Smell count | 138 |
| Workflow Guide buttons | 8 |
| VBA modules | 17 (dead-suspected 15, transformer 2) |

### 2.3 `vba-web-example.xlsm` (existing — unchanged)

| Metric | Value |
|---|---|
| Complexity score | 40 / 100 |
| Detected domain | `unknown` (none) — truthful, no business signal |
| Domain templates matched | none |
| Pillar count | 0 (no formula fan-in over threshold) |
| Anomaly count | 0 |
| Smell count | 0 |
| Workflow Guide buttons | 6 |
| VBA modules | 40 |

### 2.4 `logistics_routing_synth.xlsm` (new)

| Metric | Value |
|---|---|
| File size | 427.1 KB |
| Sheets | 6 (5 visible, 0 hidden, 1 veryHidden = `_constants`) |
| Formula count | 175 |
| Named ranges | 6 (`MAX_VEHICLES`, `SPEED_KMH`, `SERVICE_TIME_MIN`, `MAX_STOPS_PER_ROUTE`, `FUEL_PRICE`, `DRIVER_SHIFT_HOURS`) |
| Conditional formatting | 1 rule (Routes!F2:F31 red on `>1.0` capacity util) |
| Data validation | 0 |
| Complexity score | 68 / 100 (band ±5: 63–73) |
| Detected domain | `logistics-routing` (medium, 3 hits: routes, shipments, vehicles) |
| Domain templates matched | `logistics-routing` (high) |
| Pillar count | 4 (Routes!A column, Routes!C column, Shipments FUEL_PRICE, Shipments!D group) |
| Anomaly count | 0 |
| Smell count | 99 (multi-ref 50, magic-numbers 45, dup 4) |
| Workflow Guide buttons | **1** — `Run Routing Calculation` → `[0]!Execute` on Routes |
| VBA modules | 40 (donor; same distribution as capacity_planning_synth) |

**Pillar `Routes!A2..A31`**: each RouteID is referenced by 100 INDEX/MATCH
calls in `Shipments!E*` + 1 SUMIFS in `Routes!F*`, totaling fan-in 101.
Column-block dedupe collapses 30 cells into one entry.

### 2.5 `inventory_supply_synth.xlsm` (new)

| Metric | Value |
|---|---|
| File size | 428.6 KB |
| Sheets | 5 (4 visible, 0 hidden, 1 veryHidden = `_constants`) |
| Formula count | 350 |
| Named ranges | 6 (`SAFETY_FACTOR`, `ANNUAL_HOLDING_RATE`, `ORDER_COST`, `ANNUAL_DEMAND_DEFAULT`, `SERVICE_LEVEL_Z`, `LEAD_TIME_DEFAULT`) |
| Conditional formatting | 0 |
| Data validation | 1 rule (Movements!D = "in"/"out" dropdown) |
| Complexity score | 72 / 100 (band ±5: 67–77) |
| Detected domain | `inventory-supply-chain` (medium, 4 hits: movements, reorder, stock, suppliers) |
| Domain templates matched | `inventory-supply-chain` (high) |
| Pillar count | 1 (Movements!B column-block) |
| Anomaly count | 0 |
| Smell count | 106 (multi-ref 50, magic-numbers 50, dup 6) |
| Workflow Guide buttons | 0 (graceful empty) |
| VBA modules | 40 (donor) |

**Inventory-math sanity**: each row of `Stock` contains four pillar-candidate
formulas — `EOQ = SQRT(2·D·S/(H·P))`, `SafetyStock = z·σ·√LT`,
`ReorderPoint = d·LT + SafetyStock`, plus a SUMIFS aggregate from Movements
in `Reorder`. The audit's pillar detector keys on **fan-in**, not on math
significance, so the most "important" cells (EOQ formula cells) DO NOT
appear as pillars — each is referenced ≤ 1 time. Movements!B (the SKU
label column) wins by reference count.

### 2.6 `minimal_no_vba.xlsm` (new — degradation test)

| Metric | Value |
|---|---|
| File size | 6.2 KB |
| Sheets | 2 (Data, Summary; both visible) |
| Formula count | 8 (all in Summary, simple SUM/SUMIFS/AVERAGE/MAX) |
| Named ranges | 0 |
| Conditional formatting | 0 |
| Data validation | 0 |
| Complexity score | 32 / 100 (band ±5: 25–40) |
| Detected domain | `unknown` (none) — graceful |
| Domain templates matched | none |
| Pillar count | 0 (no formula reaches fan-in 20) |
| Anomaly count | 0 |
| Smell count | 30 (all multiple-references; small denominator inflates density) |
| Workflow Guide buttons | 0 (graceful "formula-driven only") |
| VBA modules | 0 (xlsm without `xl/vbaProject.bin` part) |
| H2 sections present | All required sections; **omits** "Domain-Specific Findings" (no template match — graceful conditional render) |

**Honest observation about complexity = 32/100**: the smell-density
sub-score lands at 20/20 because we have 30 multiple-reference smells
across only 132 non-empty cells — a high ratio. Without that signal the
total would be ~12. This is a known artifact of the smell-density
formula on tiny workbooks; documented but not "tuned" away because the
arithmetic is genuinely correct.

---

## 3. Cross-fixture validation outcomes

All checks passed in the Round-4 baseline run.

| Check | All 6 |
|---|---|
| audit completes (`rc == 0`) | ✓ |
| audit.md, audit.json, audit.html all written | ✓ |
| Idempotency (two runs → byte-identical md + json) | ✓ |
| Domain detection matches expectation | ✓ |
| All required H2 sections present (`Executive Summary`, `Table of Contents`, `Workflow Guide`, `Data Flow Story`, `Top Impact Findings`, `VBA Module Walkthrough`, `Reference Appendix`, `Glossary`, `Methodology`) | ✓ |
| `Domain-Specific Findings` rendered iff at least one template at high/medium | ✓ (omitted on minimal + vba-web-example, present on others) |
| Mermaid diagram count ≥ 3 in markdown | ✓ |
| `LLM-AUGMENT` markers present where applicable | ✓ |
| Zero LLM dependencies in `src/` (regex grep) | ✓ |
| Determinism (no timestamps, no absolute paths in JSON) | ✓ |

**Structure consistency assertion**: every fixture (regardless of size or
style) renders the same nine required H2 sections. Only `Domain-Specific
Findings` is conditional. Sections render with **graceful empty content**
when the underlying data is absent (e.g., minimal's "no buttons detected"
narrative in Workflow Guide).

---

## 4. Honest observations and surprises (Round 4)

1. **Word-boundary keyword bug discovered.** The original `domain.py`
   keyword list used singular forms (`route`, `vehicle`) but real
   workbooks use plural sheet names (`Routes`, `Vehicles`). The `\b...\b`
   word-boundary regex doesn't match across word characters, so
   logistics_routing_synth initially detected as `unknown`. Fixed by
   adding plural forms to `DOMAINS` in `domain.py` and `domains.py`. This
   would have affected real customers too — the synthetic corpus surfaced
   it before the friend test did.

2. **Workbook-rels path normalization bug.** `_build_sheet_id_map` in
   `workflow.py` only handled `Target="xl/worksheets/sheetN.xml"` and
   `Target="worksheets/sheetN.xml"`. The donor xlsm uses
   `Target="/xl/worksheets/sheetN.xml"` (leading slash, absolute), which
   produced corrupted paths like `xl//xl/worksheets/sheet1.xml`. Buttons
   were never detected on logistics_routing_synth until this was fixed.
   Now: handles all three forms (lstrip leading `/`, prepend `xl/` if
   missing).

3. **Minimal's complexity is 32, not the ~10 we'd intuitively guess.**
   The smell-density sub-score gets capped at 20/20 because 30 smells /
   132 cells = 227 smells per 1000 cells — a high ratio. The formula is
   working as designed; it just doesn't model "small sheet" bias. We
   chose **not** to tweak the formula to make minimal "look right" because
   that would degrade signal on real tiny operational workbooks.

4. **Capacity_planning_synth now matches `inventory-supply-chain`
   template at "low" confidence.** Round-4's plural-form additions also
   added `供应商` and `库存` to inventory keywords. Since the cap-planning
   synth has both sheets, the inventory template now scores 2 keyword
   hits → "low" confidence. The primary match remains
   `manufacturing-capacity-planning` at "high". This is a small report
   bloat (one extra Domain-Specific subsection at "low" confidence) — but
   honest. We could raise the inventory threshold from 2→3 keywords, but
   that punishes real inventory workbooks for the synth's incidental
   structure.

5. **Inventory_supply_synth landed at 429 KB, above the spec's 200-300 KB
   target.** Caused by the donor VBA blob (~330 KB). The spec explicitly
   permitted reusing the donor donut, so we accepted the size. A future
   improvement would be to use a stripped-down donor (e.g. one or two
   modules instead of 40) for variants that don't need the full VBA mass
   — would shrink to ~80 KB.

6. **Pillars in inventory_supply_synth = 1, not the ~5 a human would
   pick.** The audit picks pillars by **fan-in count**, not by
   semantic significance. The EOQ formula cells (`Stock!E*`) are
   semantically critical but each is referenced only once (in the
   Reorder sheet). Movements!B (raw SKU labels) wins by reference count
   alone. This is the expected Track-A failure mode that Tier 1.5 LLM
   augmentation will fix.

7. **Donor VBA's classification distribution is identical across all
   three donor-using fixtures** (logistics, inventory, capacity_planning):
   18 data-loader + 8 dead-suspected + 6 mixed + 5 report-writer + 3
   transformer = 40. This confirms the classifier is purely structural
   (not affected by surrounding sheets), which is the expected Track-A
   property.

---

## 5. Track A consistency conclusion

The audit pipeline is **structurally consistent** across the six-fixture
corpus:

- Same H2 skeleton on every output
- Same JSON top-level keys on every output
- Idempotent on every input
- Domain detection produces the right answer on each (after the Round-4
  word-boundary fix)
- Graceful empty paths exercised on minimal (no VBA, no named ranges,
  no buttons)
- No crashes, no traceback, no warning leaks past the renderer

What it is NOT yet able to do (deferred to Tier 1.5):
- Identify semantically-significant cells when fan-in is low
- Distinguish *intentional* domain crossover (cap-planning + inventory)
  from spurious keyword overlap
- Decompose pillar significance by formula content rather than reference
  count

These are the right things to defer — Track A's promise is "deterministic,
zero-LLM, runs anywhere" and that holds. Track B (BYOA LLM) will
specifically address #1 and #3 by reading formula text and producing
plain-language explanations.

---

## 6. How to regenerate

```bash
# Build the three new fixtures (deterministic; seed=20260509)
python tests/fixtures/build_corpus.py --variant logistics-routing \
    --out tests/fixtures/logistics_routing_synth.xlsm
python tests/fixtures/build_corpus.py --variant inventory-supply-chain \
    --out tests/fixtures/inventory_supply_synth.xlsm
python tests/fixtures/build_corpus.py --variant minimal \
    --out tests/fixtures/minimal_no_vba.xlsm

# Run audit on each
audit-xlsm tests/fixtures/logistics_routing_synth.xlsm \
    --out-dir out_logistics_routing_synth
audit-xlsm tests/fixtures/inventory_supply_synth.xlsm \
    --out-dir out_inventory_supply_synth
audit-xlsm tests/fixtures/minimal_no_vba.xlsm \
    --out-dir out_minimal_no_vba

# Run the full regression suite (51 tests, ~4 min on 3-file corpus)
pytest tests/ -q
```

The `build_corpus.py` script is **deterministic** — re-running with the
same seed produces a byte-identical xlsm (modulo openpyxl's own output,
which is deterministic on a fixed Python+openpyxl version pair). All
randomness uses `random.Random(seed=20260509)`.
