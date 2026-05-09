# Repo: tier1-audit (xlsm static auditor)

This repo ships a single-purpose CLI: `audit-xlsm`. It produces a
markdown + HTML + JSON audit of a legacy `.xlsm` workbook using pure
static analysis — pillar cells, magic-number anomalies, VBA classification,
Hermans-style code smells, complexity score. Local-only, zero LLM,
zero network at audit time.

## When the user asks to audit a workbook

The repo provides a custom agent `xlsm-audit` (see
`.github/agents/xlsm-audit.agent.md`). For multi-step audit-and-enrich
workflows, prefer invoking that agent rather than running ad-hoc
commands. It encodes the safe, verified seven-step workflow.

If the user is in a chat where the agent is not available, fall back to
running the CLI manually:

```bash
audit-xlsm <file>.xlsm --out-dir ./audit-output
```

For LLM-enriched (Track B) reports, the manual two-step is:

```bash
audit-xlsm <file>.xlsm --harness --out-dir ./audit-output
# (read prompt.md + dossier.json; produce responses.json)
audit-xlsm <file>.xlsm --ingest ./audit-output/responses.json --out-dir ./audit-output
```

## Architectural promise — never violate

This tool's value proposition includes that **the package itself never
calls an LLM**. There is no `requests`, `httpx`, `anthropic`, or `openai`
import in the source tree, and there are tests that guard this fact
(`tests/test_corpus.py::test_zero_llm_in_source`,
`tests/test_harness.py::test_harness_zero_llm_imports`). When suggesting
edits, never propose adding an LLM client library; the BYOA model is
load-bearing for compliance.

LLM-driven enrichment happens in the user's existing chat tool (Copilot
Chat / Claude Desktop / ChatGPT). When you, the LLM, are asked to fill
narratives via the `xlsm-audit` agent, you are the BYOA component —
your tokens, the user's session, no API call from our package.

## Coding conventions

- Python 3.9+ baseline. No f-string features that require 3.10+.
- Stdlib + the three pinned deps (`openpyxl`, `oletools`, `formulas`)
  only. No new dependencies without explicit approval.
- Tests live under `tests/`; the corpus test (`tests/test_corpus.py`)
  is the integration floor — 65/65 must stay passing.
- Code is in `src/tier1_audit/` as a flat package; one module per
  concern (`audit.py`, `harness.py`, `ingest.py`, `pillars.py`, …).

## When suggesting edits

Treat the report rendering (`render.py`) and the audit pipeline
(`audit.py`) as stable surfaces — changes to either ripple through
the corpus tests. Prefer additive edits in new modules over rewriting
the existing ones.
