# Output sample

This is a curated excerpt from a real run of `audit-xlsm` on a synthetic
capacity-planning workbook (1 file, ~15K cells, ~4.5K formulas, 40 VBA
modules, 12 sheets including 2 hidden + 2 very-hidden). It is intended to
show what the report looks like end-to-end — every section is real output,
not mocked.

## Section 1-2 — basics

```markdown
# Audit report — `capacity_planning_synth.xlsm`

> Tier 1 audit. Pure static analysis — no AI, no Excel, no macro execution.

## 1. File metadata

| Field | Value |
|---|---|
| File name | `capacity_planning_synth.xlsm` |
| File size | 540,479 bytes (527.8 KB) |
| SHA-256 | `47df0977768fa5031eb2f31df1bbf8a3461857384ff7012777f13a9016bed205` |

## 2. Basic statistics

| Metric | Value |
|---|---|
| Sheet count (total) | 12 |
| Sheet count visible / hidden / veryHidden | 8 / 2 / 2 |
| Non-empty cells | 15,551 |
| Formula cells | 4,585 |
| Named ranges | 6 |
| VBA modules | 40 |
| VBA total lines | 14,709 |
```

The 2 hidden + 2 very-hidden sheets are the first hint that the workbook
is hiding state from casual readers — both flagged in the risk section
later.

## Section 5 — complexity score

```markdown
## 5. Complexity score

**Total: 87 / 100**

| Sub-score | Value | Rationale |
|---|---|---|
| data scale | 24/20 (capped at 20) | log10(15551) cells -> 20/20 |
| formula depth | 20/20 (capped at 20) | max-cond=4, max-ops=20, max-chain=0 -> 20/20 |
| vba mass | 20/20 (capped at 20) | 14709 LOC across 40 modules -> 20/20 |
| smell density | 20/20 | 154 smells / 15551 cells -> 20/20 |
| metadata complexity | 11/20 | 2 hidden + 2 veryHidden + 6 named ranges + 154 cross-sheet refs -> 11/20 |
```

87 / 100 is the ceiling for a workbook that maxes out 4 of 5 dimensions.
Only metadata is moderate — 6 named ranges is small.

## Section 6.5 — Pillar cells (the headline logic-depth feature)

This is one of the three new Tier 1 logic-comprehension modules.

```markdown
## 6.5 Pillar cells — systemic single-points-of-impact

_Cells with the highest fan-in (most-referenced). Modifying any of these
cascades through many formulas; treat them as critical change points._

| Rank | Cell | Fan-in | Affected sheets | Kind | Confidence | Narrative |
|---|---|---|---|---|---|---|
| 1 | `_constants!C4` | 80 | `BOM`, `MPS`, `辅助计算` | constant-input | high | Modifying `_constants!C4` (a constant input) would cascade into 80 formulas across sheets `BOM`, `MPS`, `辅助计算` — high-risk change point. |
| 2 | `订单!C10` | 72 | `MPS` | constant-input | high | Modifying `订单!C10` (a constant input) would cascade into 72 formulas across sheet `MPS` — high-risk change point. |
| 3 | `订单!C100` | 72 | `MPS` | constant-input | high | Modifying `订单!C100` (a constant input) would cascade into 72 formulas across sheet `MPS` — high-risk change point. |
| ... | ... | ... | ... | ... | ... | ... |

**Top-5 pillar drilldown — sample dependents:**

- **`_constants!C4`** (80 dependents):
    - `BOM!C22`
    - `BOM!C23`
    - `BOM!C24`
    - `BOM!C25`
    - `MPS!H44`
```

Reading this as a BA: one cell — `_constants!C4`, an unassuming row in a
hidden constants sheet — drives 80 formulas across three other sheets.
That's exactly the "if I touch this it could blow up the whole MPS sheet"
warning that legacy-Excel maintainers normally have to discover the hard
way.

The 19 entries in `订单!C*` (the order list column) are correct but
repetitive — they reflect symmetric design (every order row is referenced
the same way). A future Tier 2 might collapse these into a "column-level
pillar" entry.

## Section 6.6 — Magic-number anomalies (synthetic fixture)

The synth corpus uses parameter-sweep formulas where every cell uses a
different value, so nothing flags as an anomaly. To demonstrate detection,
this excerpt is from a tiny synthetic fixture (`tests/fixtures/anomaly_fixture.xlsm`):
29 cells with discount formulas like `=C<n>*0.85`, except row 27 deliberately
uses `0.82`.

```markdown
## 6.6 Magic-number anomalies — outliers within formula clusters

_Inside groups of cells that share the same formula shape, this section flags
positions where a small minority uses a different numeric constant. These are
exactly the kind of "why is this row's discount different?" findings that
often indicate a missed update or a deliberate (but undocumented) carve-out._

| # | Cluster sample | Cluster size | Mode value | Outlier value | Outlier locations | Confidence | Narrative |
|---|---|---|---|---|---|---|---|
| 1 | `=C2*0.85` | 29 | `0.85` (28/29) | `0.82` (1/29) | `Pricing!D27` | high | In a cluster of 29 formulas matching `=C2*0.85`, position #1 usually equals `0.85` (28/29 = 97%), but only `Pricing!D27` use `0.82` (deviation 0.03). Confirm whether this is intentional. |
```

When this kind of pattern actually exists in a workbook, it's the kind of
finding that makes a BA call the original author and ask "wait, was row 27
supposed to use 0.82?" — usually it's a missed update from a rate-change
that never propagated. Sometimes it's a deliberate carve-out that nobody
documented.

The Stage-1 audit corpus (3 real-world legacy workbooks) does not exhibit
this pattern — they happen to be parameter sweeps, not "everyone-uses-the-
same-rate" tables. We document this honestly: 0 hits across 3 files is a
real finding, not a bug.

## Section 8 — VBA module classification (sample)

```markdown
## 8. VBA modules

| Module | Type | LOC | #Sub | #Func | Inferred type | Confidence | Reads | Writes | Ext calls | OnErrorResumeNext |
|---|---|---|---|---|---|---|---|---|---|---|
| `Module1.bas` | standard | 68 | 2 | 0 | **transformer** | low | — | — | no | no |
| `Sheet1.cls` | class_or_document | 8 | 0 | 0 | **dead-suspected** | medium | — | — | no | no |
| `Salesforce.bas` | standard | 125 | 0 | 7 | **transformer** | medium | — | — | no | no |
| `Gmail.bas` | standard | 87 | 1 | 2 | **data-loader** | high | — | — | no | no |
| `OPSSheet.cls` | class_or_document | 55 | 3 | 0 | **report-writer** | low | — | — | no | no |
| `WindowsAuthenticator.cls` | class_or_document | 79 | 4 | 0 | **dead-suspected** | medium | — | — | yes | no |
| `WebHelpers.bas` | standard | 3177 | 11 | 53 | **data-loader** | medium | — | — | yes | yes |

**VBA module details — external keywords & range literals:**

| Module | External keywords | Range literals (uniq) | Classifier rationale |
|---|---|---|---|
| `Module1.bas` | — | 6 | name uninformative; reads=17, writes=10, loops=7 |
| `Sheet1.cls` | — | 0 | empty/near-empty shell: 0 subs, 0 calls, 0 writes, 8 LOC |
| `Salesforce.bas` | — | 0 | name pattern vote: transformer=2 |
| `Gmail.bas` | — | 0 | name pattern vote: data-loader=2 |
| `WebHelpers.bas` | WinHttpRequest | 12 | name pattern vote: data-loader=4, others={data-loader: 4, transformer: 1, report-writer: 2} |
```

What this tells a BA at a glance:
- 16 of 17 OperationPlanning modules are empty `Sheet*.cls` shells —
  `dead-suspected`. The actual logic lives in `Module1` (transformer) and
  `Module2` (transformer).
- In a 40-module library imported from open-source vba-web, most modules
  are correctly recognized as data-loaders / transformers / dead.
- `WebHelpers.bas` has a `WinHttpRequest` external keyword — anyone
  worried about the workbook making outbound calls knows exactly where to
  look.

## Section 9 — Risk indicators

```markdown
## 9. Risk indicators

| Indicator | Value |
|---|---|
| Hidden sheets | 2 (`_constants, Forecast P3`) |
| Very-hidden sheets | 2 (`_audit, _params`) |
| Cross-sheet referencing formulas | 154 |
| Cells with formula errors (cached) | 0 |
| External workbook reference patterns | 0 |
| Circular reference suspects | 0 |
| Parse errors logged | 0 |
```

The four hidden / very-hidden sheets named with leading underscores
(`_constants`, `_audit`, `_params`) are a tell-tale for "infrastructure
sheets the original author kept out of the way" — inspecting them directly
is usually the first thing a forensics-minded BA does. The audit surfaces
their existence; reading their contents is a manual follow-up step.
