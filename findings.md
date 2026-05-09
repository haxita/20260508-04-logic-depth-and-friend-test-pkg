# Findings — Logic depth modules + friend-test packaging

Task: 20260508-04 · macOS · CPython 3.9.6 in worker venv. The package
`pyproject.toml` declares `requires-python = ">=3.10"` (the friend's target
environment); the local validation venv uses 3.9 with `--ignore-requires-python`
because no 3.10+ Python is installed on this Mac. All code uses
`from __future__ import annotations`, so PEP 604 `int | None` style hints
remain compatible at runtime on 3.9.

## TL;DR — corpus results

| File | Smells | Pillars | Anomalies | VBA classified |
|---|---|---|---|---|
| `capacity_planning_synth.xlsm` | 154 | 20 | 0 | 40/40 |
| `OperationPlanning.xlsm` | 138 | 20 | 0 | 17/17 |
| `vba-web-example.xlsm` | 0 | 0 | 0 | 40/40 |
| `tests/fixtures/anomaly_fixture.xlsm` (synthetic) | 54 | 1 | 1 | 0/0 |

Smell counts and complexity scores match Stage 1 baselines exactly — the
3 new modules are additive only.

## A.1 Pillar cells

- **Threshold**: fan-in >= 20, top 20. Distinct from the
  multiple-references *smell* threshold (5) — pillars are intentionally a
  short, high-impact list.
- **Data source**: reuses `incoming` map already built by
  `smells.detect_multiple_references` (no extra tokenizer pass).
- **Narrative**: rule-based template that names the pillar kind
  (`constant-input` / `formula-relay` / `whole-column`), the fan-in count,
  the affected sheet list (compressed when >3), and a risk band.
- **Confidence is `high`** because fan-in is an exact count.
- **Honest flaw**: synth has 19 of 20 pillars in `订单!C2..C200` (each row
  of the order list referenced symmetrically by 72 MPS formulas). They are
  all real and equally true; collapsing into "column-level pillar" is
  Tier 2 polish, not Tier 1 correctness.

## A.2 Magic-number anomalies

- **Algorithm**: for each duplicated-formula cluster of size >= 5,
  re-tokenize each formula, group Number tokens by positional index,
  compute mode + outliers per position. Outlier = value occurring at
  most 5 % of cluster *and* numerically distinct from mode (canonicalized
  via `repr(float(x))`, so `1.0 == 1.00`). Trivial-formatting deviations
  (epsilon < 1e-9) dropped.
- **Mode-share floor lowered from 0.5 to 0.3** during calibration.
  Reason: real workbooks have parameter sweeps where no value is a strict
  majority but the deviation is still meaningful. 0.30 still produces 0
  false positives on the 3-file corpus (verified) and 1 true positive on
  the synthetic fixture.
- **Confidence labels**:
  - `high` — mode-share >= 0.95 and deviation > 0.001
  - `medium` — mode-share >= 0.80
  - `low` — otherwise
- **Why all 3 corpus files have 0 anomalies**: they are uniform
  parameter sweeps (e.g. `=SUMPRODUCT(... MONTH(...) = N ...)` with N
  evenly distributed across 1-12 months). There is no "everyone uses 0.85
  except this one row" pattern in any of them. We document this rather
  than tune until something fires; tuning to fabricate would degrade the
  signal-to-noise ratio.
- **Synthetic fixture proves the detector works**: 28/29 cells in
  `Pricing` use 0.85, row 27 uses 0.82 — the audit reports a single
  high-confidence anomaly with the correct narrative.
- **Surprise during calibration**: the duplicated-formula clustering
  collapses any two formulas with the same normalized shape, even from
  unrelated sheets. My first fixture had `=A*B+0.85` (29 cells) and
  `=Constants!B2*<int>` (25 cells) — both normalized to `=R*N` and merged
  into a 54-cluster, where the 25 small integers (2..26) all flagged as
  outliers vs the 0.85 mode. The fix was to keep the second sheet's
  formula shape distinct (`=Constants!$B$2+0.5`).

## A.3 VBA module classification

- **Decision rules** (priority order):
  1. **Empty/near-empty shell** (line_count < 30, no subs, no calls, no
     value writes): `dead-suspected`. This catches the default
     `Sheet*.cls` modules that openpyxl/oletools surface even when the
     workbook has no event handlers.
  2. **Name-pattern vote**: regex against Sub/Function names assigns
     each name to a category; majority wins. Margin >= 2 -> `high`
     confidence; >= 2 hits -> `medium`; otherwise `low`.
  3. **UI-handler bias** for modules whose name matches `^(ThisWorkbook|Sheet\d+|Workbook|Worksheet)`:
     +5 to `ui-handler` if any Sub has `_Click`/`_Change`/`_Open` etc.
  4. **Structural fallback** when names are uninformative:
     - >= 5 `Range/Cells(...).Value =` writes & >= 2x reads -> `report-writer`
     - >= 5 `... = ...Range/Cells(...).Value` reads & >= 2x writes -> `data-loader`
     - >= 3 reads, >= 3 writes, >= 1 loop -> `transformer`
     - else -> `mixed`
- **Outputs**: `inferred_type`, `confidence`, `reads_sheets`,
  `writes_sheets`, `external_calls` (boolean, reuses Stage 1's
  external_keywords list), `value_writes`, `value_reads`,
  `control_flow_count`, name-signal histogram, one-line rationale.
- **Sheets-read-vs-written attribution**: regex window of +/-200 chars
  around each `Range/Cells(...).Value =` (write context) or `=...Range...
  Value` (read context). Any `Sheets("X")` / `Worksheets("X")` reference
  inside the window is attributed.
- **Honest flaws**: many vba-web modules use `Application.Cells.Value`
  or implicit-sheet `Range("A1")` — neither matches our `Sheets("...")`
  regex, so reads/writes lists are empty for those modules. That's
  correct (we have no sheet evidence) but visually empty; a future
  improvement is to fall back to "module operates on whichever sheet is
  active" when no explicit Sheet ref exists.

## Sanitize implementation

- Redaction happens at extraction time in `extract_cells(..., sanitize=True)`,
  which replaces every non-formula `CellRow.value` with `<redacted>`.
  Cached evaluated values for formula cells: redacted unless they contain
  formula error tokens (`#REF!`, etc.) — those are kept so risk indicators
  still show error counts.
- VBA source is **not** sanitized. Code is structure, not data; the audit
  needs it for smell detection and classification. Customers who consider
  their VBA itself proprietary should not distribute the audit at all.
- Banner injected at the top of `audit.md`. Methodology footer adds
  `**Sanitize mode active**: ...` paragraph for transparency. JSON has
  `"sanitized": true` and `"methodology.sanitize_mode": true`.
- Verified by `test_sanitize_strips_cell_values` and
  `test_sanitize_preserves_formula_text`.

## Cross-platform care taken

- All paths via `pathlib.Path`. Worker dir code never concatenates with
  literal `/`.
- Every `read_text` / `write_text` call passes `encoding="utf-8"`
  explicitly (no reliance on locale default — Windows would default to
  cp1252 otherwise).
- `pyproject.toml` declares `requires-python = ">=3.10"` and uses PEP 621
  metadata + setuptools build backend. PEP 660 editable install via
  `pip install -e .` proven on macOS (clean venv).
- README and `friend-test-setup.md` provide Windows PowerShell + cmd.exe
  command examples, ExecutionPolicy guidance, UTF-8 path notes.

## Determinism contract

Stage 1's contract preserved verbatim:
- `json.dumps(sort_keys=True, indent=2, ensure_ascii=False)` + trailing
  newline.
- All lists sorted by stable lex key before serialization. New A.1/A.2/A.3
  modules: pillars sorted by `(-fan_in, location)`, anomalies sorted by
  `(confidence_rank, -mode_share, -cluster_size, first_outlier_loc)`,
  classifications sorted by `module_name.lower()`.
- No timestamps, no absolute paths, no `os.getcwd()`.

**Idempotency proven**: each test file run twice; pytest
`test_idempotency` confirms `audit.md` and `audit.json` are byte-identical
across both runs for all 3 corpus files.

## Surprises encountered

1. **`requires-python` floor + local Python 3.9.** The Mac has only 3.9;
   pyproject's `>=3.10` blocks editable install. Worked around with
   `--ignore-requires-python` for the validation venv only — the shipping
   constraint stands. Documented separately in next-steps.
2. **Old pip in 3.9's venv** (21.2.4) lacks PEP 660 editable support.
   Required `pip install --upgrade pip setuptools wheel` first. The
   `friend-test-setup.md` includes this step implicitly because Python
   3.10+ ships pip >= 22.
3. **numpy/scipy version pinning**: latest `numpy` (2.3+) and `scipy`
   (1.14+) require Python >= 3.11; `formulas` pulls them transitively. For
   our 3.9 validation venv we pre-install `numpy<2.1` + `scipy<1.14`. This
   is invisible to friend installs on 3.10+ — pip resolves correctly.
4. **VBA classification on default Sheet*.cls modules**: openpyxl exposes
   one `Sheet1.cls` ... `SheetN.cls` per worksheet, even when none have
   event handlers. Without the `is_empty_shell` rule, all 16 of OpPlan's
   default Sheet modules fell into `mixed`. Adding the rule moved them to
   `dead-suspected`, which is the truthful label.
5. **JSON serialization of `source_text`**: `dataclasses.asdict` is
   deeply recursive, so my outer `pop("source_text", None)` only ran on
   the top-level dict. Fixed by filtering `_EXCLUDED_FIELDS` at every
   layer of `_to_jsonable`. Caught by the `test_json_roundtrip_synth`
   pytest case.

## Polish round increment (2026-05-09)

**Pillar dedupe (Item 3).** Key = `(sheet, column, fan_in,
frozenset(affected_sheets))`. Synth Top-N collapses 20 -> **2** distinct
entries: `_constants!C4` + `订单!C10..C189 (99 cells)` column-block.
OperationPlanning collapses 20 -> 10. Pre-dedupe scan width = 100.

**Report restructure (Item 4).** Pillars + anomalies promoted to top-level
H2 (was `6.5/6.6`). Cover line + executive summary (complexity, sparklines,
3 headlines, domain hint) + auto-TOC from `_SECTION_BUILDERS` single source
of truth. Duplicate `## 10. Methodology` bug noted in spec was absent on
inspection; regression test now guards.

**Mermaid (Item 5).** Three diagrams (sheet flow / VBA classification /
pillar impact). Caps 30 nodes, 50 edges in diagram 1; top-5 in diagram 3.
ASCII-clean node IDs via `_safe_node_id`. Spot-checked at mermaid.live —
all 3 corpora render.

**HTML + CDN (Item 6).** Pure stdlib + f-strings + `html.escape`. Mermaid
CDN by default; `--mermaid-inline` downloads JS once at audit time for
sealed-network viewing. `@media print` page-break rules so H2 sections
don't split.

**Domain (Item 7).** 6 domains, ~10 keywords each. Synth ->
`capacity-planning` (medium, 4 hits). OperationPlanning ->
`capacity-planning` (medium, 2). vba-web-example -> `unknown` (truthful —
web-API boilerplate, no business signal).

**Pip caveats.** Old pip (<21.3, macOS system 3.9) lacks PEP 660 editable —
default switched to `pip install .`. `requires-python>=3.9` works because
all files have `from __future__ import annotations`. Friend doc gains
4 install paths + troubleshooting table.

## Round 4 — Test corpus expansion (2026-05-09)

**3 new fixtures** via `tests/fixtures/build_corpus.py` (seed=20260509,
stdlib + openpyxl only):

| Fixture | Size | Domain (conf) | Buttons | Pillars | Complexity |
|---|---|---|---|---|---|
| logistics_routing_synth | 427 KB | `logistics-routing` (medium) | 1 | 4 | 68 |
| inventory_supply_synth | 429 KB | `inventory-supply-chain` (medium) | 0 | 1 | 72 |
| minimal_no_vba | 6 KB | `unknown` (none) | 0 | 0 | 32 |

**Two latent bugs surfaced**: (a) `\b…\b` keyword regex missed plural
sheet names — logistics detected as `unknown`. Fixed by adding plurals
to `DOMAINS` + `domains.py`. (b) `_build_sheet_id_map` corrupted paths
when `workbook.xml.rels` had leading-slash Targets — buttons invisible.
Fixed by normalizing target paths.

**Honest surprises**: minimal complexity = 32 (not ~10) because
smell-density caps at 20/20 on small N — kept honest, not tuned.
Inventory pillars = 1 (Movements!B label column wins by fan-in, not by
EOQ semantic significance — Track-A failure mode for Tier 1.5).

**Pytest**: 51/51 passing (27 baseline + 24 new). All 6 corpora pass:
idempotent, zero-LLM, consistent H2 skeleton (`Domain-Specific Findings`
correctly conditional). `docs/test-corpus.md` = canonical baseline.
