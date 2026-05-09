# Next steps — from this beta to a friend-validated v0.1, then onward

Audience: GM Assistant + Michael, deciding what to build / decide / ship next.

## Tier 1.5 LLM-augmented hook design (the next-stage interface)

The product north star puts Tier 1.5 above Tier 1 as "LLM-assisted semantic
explanation, BYOA". Where does the LLM enter the existing pipeline?

**The natural insertion point is post-`build_audit`, pre-`render_markdown`.**
The `AuditReport` dataclass is the contract; the LLM consumes it and emits
augmented narratives that the renderer interpolates into existing sections.

Concrete proposal:

```python
# tier1_audit/llm.py (Tier 1.5, NOT in 0.1.0)
class LLMAugmenter(Protocol):
    def explain_pillar(self, p: Pillar, context: AuditReport) -> str: ...
    def explain_anomaly(self, a: MagicNumberAnomaly, context: AuditReport) -> str: ...
    def summarize_vba_module(self, vm: VbaModule, classification: VbaClassification) -> str: ...

def augment(report: AuditReport, llm: LLMAugmenter) -> AuditReport:
    """Run LLM passes; mutate report in-place to add `narrative_llm` fields."""
```

The renderer would prefer `narrative_llm` when present, fall back to the
zero-LLM `narrative` otherwise. Tier 1 stays free + reliable; Tier 1.5
adds depth without rewriting.

**BYOA via `gh copilot explain`**: the simplest BYOA implementation is to
shell out to `gh copilot explain "..."` (subprocess). Customer's Copilot
license + corp account does the work; we ship zero AI dependencies. Same
approach for Claude CLI: shell out to `claude -p "..."`. The
`LLMAugmenter` Protocol abstracts the choice. We ship two reference
implementations (`CopilotCLIAugmenter`, `ClaudeCLIAugmenter`) plus a
`MockAugmenter` for tests.

**What the LLM is asked to do**:
- For pillars: read the formula text of the top dependents and write a
  one-paragraph "this cell controls the X calculation in the Y workflow"
  explanation.
- For anomalies: take the cluster sample formula + outlier location +
  surrounding cells and explain "this looks like a forgotten rate update"
  vs. "this looks like an intentional carve-out — confirm with author".
- For VBA modules: take the source + classification label and produce a
  3-sentence "what this module does" summary in plain English.

**Cost guardrails**: each LLM call is small (< 2K input tokens). With 20
pillars + 30 anomalies + 40 modules ≈ 90 calls per workbook. At Copilot
CLI prices (BYOA — paid by the customer's existing license), free.

## VS Code extension wrapping path (Stage 4)

The product north star says VS Code extension is the primary
distribution. The Stage 4 wrapper should:

1. **Bundle the `tier1_audit` Python package** as the analysis engine.
   Extension is TS/JS UI; Python is the worker. (Same model as
   Microsoft's own Python extension delegates to Python interpreter
   processes.)
2. **Tree view**: "Audit your workbook" command opens a file picker;
   audit runs in background; results render as a webview panel and a
   tree of findings.
3. **Click-through navigation**: clicking a pillar cell or smell finding
   opens the actual `.xlsm` (well — opens a placeholder "click to view in
   Excel" since VS Code doesn't render xlsm; or opens the formula text
   in a virtual editor tab).
4. **`--sanitize` toggle in the UI**: a checkbox in the audit panel.
5. **BYOA configuration**: settings.json field
   `tier1Audit.llmProvider: "copilot-cli" | "claude-cli" | "none"`.
   Defaults to none → free Tier 1 only.
6. **Marketplace packaging**: VSIX bundle, signed with a publisher key
   Michael owns. The Marketplace listing includes the same compliance
   talking points from the README.

Stage 4 is downstream — Stage 2's friend test must succeed first. But
keeping the package boundary clean now (CLI as the entry point, dataclass
report as the data model) means the Stage 4 wrap is mostly UI work, not
re-architecting.

## Expected friend feedback handling flow

The friend will likely report one of four things:

1. **"It worked, here's the report excerpt"** → Michael reviews, scores
   subjective quality of pillars / anomalies / VBA labels. If positive,
   the friend test is conclusive; we move to Tier 1.5 work.
2. **"It crashed on my file"** → reproduction without the file is hard.
   Best ask: send the `audit-failed.md` traceback (sanitize the file path
   if needed). 9 of 10 crashes will be (a) malformed VBA we can't parse,
   (b) cell-level data the openpyxl version chokes on, (c) Windows
   path/encoding edge case. Add a regression test against the
   reproduction; ship 0.1.1.
3. **"The labels are wrong"** → the most valuable feedback. Ask which
   labels feel wrong + why. Tune the heuristic name patterns; rebuild
   the synth corpus to lock the change in. Categorize feedback as:
   *false positive* (over-eager flag) vs *false negative* (missed real
   issue) vs *taxonomy mismatch* (label correct but customer expected
   a different vocabulary).
4. **"My company won't let me install random packages"** → expected. The
   `friend-test-setup.md` compliance section is the first response. If
   IT says no, ask what they'd accept (e.g. an internally-mirrored PyPI,
   a signed wheel via their vendor portal). Don't push — this becomes
   an early data point about enterprise distribution friction. The
   long-term answer is the VS Code extension via Microsoft's vetted
   channel.

## What's deferred from spec

- **Tier 1.5 LLM-augmented depth**: interface designed above; not built.
- **VS Code extension**: Stage 4.
- **HTML / PDF rendering**: Tier 2.
- **Multi-file batch + corpus dashboard**: Tier 1 v1 (post-friend-test).
- **GitHub repo creation + push**: Michael decides public/private and
  triggers manually.
- **Pillar column-level dedup**: Tier 2 polish — when 19 of 20 pillars
  are `订单!C*`, collapse into a single "column 订单!C is a pillar"
  entry. Saves report space, doesn't change the underlying analysis.

## Risks for the friend test specifically

1. **Local Python version mismatch.** We declared `>=3.10`; if the
   friend has 3.9, install will fail with a clear pip error. The
   compliance doc tells them what to ask IT for.
2. **xlsm with Power Query / connections**: not parsed. Audit will
   complete but those features will be invisible. Document as known
   limitation.
3. **Encrypted xlsm**: oletools may surface a different error path.
   Untested in the corpus. If the friend hits this, we add a clear
   "this file is password-protected" message.
4. **Performance on a 50K-cell workbook**: the synth has 15K cells / ~1
   minute audit. Linear scaling would be ~3 min for 50K. Token-pass +
   range expansion is the bottleneck. If the friend's file is much
   larger, we may need to add `--quick` mode that skips the chain
   analysis (already capped at 20K formulas; just makes the cap visible
   to the user).

## What to do this week

1. Michael creates the GitHub repo (public or private — his call) and
   pushes the `workers/20260508-04-...` contents (excluding `.venv*` and
   `out_*` directories — `.gitignore` handles this).
2. Friend gets the repo URL + the `friend-test-setup.md` link.
3. Wait 1–2 weeks for feedback.
4. Triage feedback into the 4 buckets above.
5. Decide whether 0.1.x patches are needed or whether to start Tier 1.5
   architecture.

## Polish round forward-looking notes (2026-05-09)

- **Domain dictionary**: 6 domains × ~10 keywords seeded in
  `src/tier1_audit/domain.py`. Grow per friend feedback; avoid pre-emptive
  bloat.
- **Mermaid caps**: 30 nodes / 50 edges in diagram 1. Real 100-sheet
  workbooks will truncate; `+ N more` is the visible signal. Follow-up
  could expose `report._sheet_edges` as `audit.dot` for external tooling.
- **`--mermaid-inline` failure mode**: CDN fetch fails -> WARN + fall back
  to CDN. Sealed-network friends need either an internally-mirrored JS
  bundle or to render on a connected laptop and ship the HTML.
- **Pillar dedupe boundary**: group key requires *identical* fan-in. A
  100th cell with fan-in=73 (vs the 99 with fan-in=72) becomes its own
  entry. Keep strict — softening with tolerance loses the "same column
  block" assertion. Re-evaluate if friends report noise.
- **Tier 2 overlay** (out of scope): exec-summary findings are the seed
  for governance scores ("change-blast-radius: 80 formulas, hidden-logic
  load: 2 veryHidden sheets"). Data already in `report.*`; Tier 2 is
  presentation, not new analysis.
