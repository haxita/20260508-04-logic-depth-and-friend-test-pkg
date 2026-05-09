# tier1-audit

> Static-analysis audit for legacy `.xlsm` workbooks — pillar cells,
> magic-number anomalies, VBA module classification.
> **Local-only. Zero LLM. Zero network.**

## What

`tier1-audit` reads a single `.xlsm` workbook and produces a deterministic
markdown + JSON report that explains:

- **What's in it** — sheets, formulas, named ranges, conditional formatting,
  data validation, VBA modules.
- **What's risky** — Hermans-style code smells (multiple-references,
  long-calculation-chain, conditional-complexity, multiple-operations,
  magic-numbers, duplicated-formulas), formula errors, hidden sheets,
  external workbook references, circular reference suspects.
- **What the logic actually does**:
  - **Pillar cells** — every cell ranked by fan-in, with a plain-English
    "modifying X cascades into Y formulas across sheets A/B/C" narrative.
  - **Magic-number anomalies** — inside groups of cells that share the same
    formula shape, this section flags positions where a small minority uses
    a different numeric constant. ("Why is row 27's discount 0.82 instead of
    0.85 like everyone else?")
  - **VBA module classification** — every module gets a heuristic label:
    `data-loader` / `transformer` / `report-writer` / `ui-handler` /
    `dead-suspected` / `mixed`, plus the sheets it reads from / writes to,
    and whether it makes external/COM calls.
- **A complexity score** 0-100, broken into five 0-20 sub-scores:
  data scale, formula depth, VBA mass, smell density, metadata complexity.

This is the **Tier 1** capability of a tiered offering. Tier 1 is structural
and statistical only — no AI, no semantic interpretation. The audit
deliberately stops at "here are the facts, ranked"; humans decide what to
do with them. (Future Tier 1.5 adds LLM-augmented semantic explanation
under a Bring-Your-Own-Account model — your Copilot or Claude CLI, not
ours.)

## Quick start

Requires **Python 3.9 or newer**.

```bash
pip install .
audit-xlsm path/to/your-file.xlsm
open out/audit.md          # macOS — or use any markdown viewer
open out/audit.html        # browser-rendered version with diagrams
```

Three commands. The audit writes `audit.md`, `audit.html`, and `audit.json`
into `./out/` by default; pass `--out-dir <dir>` to choose another location.

Pick output formats with `--format`:

```bash
audit-xlsm file.xlsm --format md     # markdown only
audit-xlsm file.xlsm --format html   # html only
audit-xlsm file.xlsm --format both   # default
```

Want fully offline HTML (no Mermaid CDN at view time)?

```bash
audit-xlsm file.xlsm --mermaid-inline
```

For corporate / sealed-network installs, see `docs/friend-test-setup.md`
which covers four install paths (default `pip install .`, GitHub direct,
internal Artifactory/Nexus mirror, corporate proxy).

## Privacy

```bash
audit-xlsm path/to/your-file.xlsm --sanitize
```

`--sanitize` redacts every non-formula cell value in the report. Formulas,
VBA source, smells, structural counts, and the SHA-256 of the original file
are preserved. The audit.md gets a banner at the top:

> 🔒 SANITIZED MODE — no cell values in this report.

**Recommended workflow if you're considering sharing the report:**
1. Run with `--sanitize` first.
2. Read the report.
3. Confirm nothing sensitive is present.
4. Decide whether to share that report (or a redacted excerpt of it).

The tool runs entirely on your local machine. No data is sent anywhere.
There are no telemetry calls, no API keys, no network requirements.

## What's in the report

| Section | Content |
|---|---|
| 1. File metadata | name, size, SHA-256 |
| 2. Basic statistics | sheet/cell/formula/named-range/CF/DV/VBA counts |
| 3. Sheets | per-sheet rows × cols × non-empty × formulas × max ref |
| 4. Named ranges | name, scope (workbook or sheet), reference |
| 5. Complexity score | total + 5 sub-scores + rationale |
| 6. Smells | 6 Hermans-style smell families with severity + confidence |
| **6.5 Pillar cells** | **top fan-in cells with narrative** |
| **6.6 Magic-number anomalies** | **outliers in formula clusters** |
| 7. Magic-number index | top 20 non-trivial numeric literals |
| 8. VBA modules | per-module LOC, subs, **inferred type**, reads/writes |
| 9. Risk indicators | hidden sheets, errors, external refs, cycles |
| 10. Methodology | library versions, thresholds, semantics |

A real-world output excerpt is in `docs/output-sample.md`.

## Compliance note

This tool runs entirely on your local machine, makes no network calls, and
requires no LLM API access. It is statically-analyzed pure Python. The
runtime dependencies are three open-source libraries:

- [`openpyxl`](https://pypi.org/project/openpyxl/) (MIT license) — reads
  cell + structural data from `.xlsm`. ~10M downloads/month.
- [`oletools`](https://pypi.org/project/oletools/) (BSD license) — extracts
  VBA module source from the OLE container. Used by SOC analysts and
  malware researchers; well-known in security tooling.
- [`formulas`](https://pypi.org/project/formulas/) (EUPL-1.1) — Excel
  formula tokenizer. **We use the tokenizer only — no formula evaluation.**

No macro is ever executed. No formula is ever evaluated. The audit reads
file contents, tokenizes formulas, parses VBA text — all statically.

If your IT department wants to verify, the Python source is open in `src/`
(< 2,500 lines total) and easy to inspect. Running with `--sanitize` and
inspecting the report before sharing is the standard mitigation.

## Status

**BETA.** Known limitations:

- xlsm-first by design. `.xlsb` is not supported natively (Save As → xlsm
  in Excel; ~30 seconds).
- Whole-column references (`SUM(A:A)`) are treated as a single opaque ref;
  cells under such refs are under-counted in fan-in.
- Tiled formulas (the same formula dragged across thousands of cells) show
  up correctly as a duplicated-formulas finding but are not separately
  distinguished from "scatter" duplication.
- VBA classification is heuristic — a "transformer" label is a structural
  guess, not a semantic claim.
- Multi-file batch + corpus-level dashboards are not in this version.

See `next-steps.md` for the roadmap to v1.

## License

Apache License, Version 2.0. See `LICENSE`.

## Versioning

Version 0.1.0. Pre-1.0 — minor versions may break command-line behavior
or report layout. Within 0.1.x, behavior is stable.
