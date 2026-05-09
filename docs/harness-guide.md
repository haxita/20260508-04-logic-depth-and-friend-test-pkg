# Track B (BYOA) Harness — End-to-End Guide

## What this is

`tier1-audit` ships with two output modes:

- **Track A (default)**: pure static analysis. Outputs `audit.md`, `audit.html`,
  `audit.json`. Heuristic narratives only — what the code structurally does,
  not what it MEANS for the business. Zero LLM. Zero network.
- **Track B (this guide)**: BYOA — Bring Your Own AI. The tool emits a
  `dossier.json` + `prompt.md` package you paste into your own LLM client
  (VS Code Copilot Chat, Claude Desktop, ChatGPT web, claude CLI — any of
  them). The LLM returns a JSON object with semantic narratives. You save
  it as `responses.json` and run a second command to splice the narratives
  into the audit. The result is `audit-enriched.md` + `audit-enriched.html`.

**The architectural promise**: this tool itself never calls an LLM. It only
prepares text for you to paste into a tool you already use. Your keystrokes
into Copilot Chat are YOUR keyboard activity — not a network call from this
tool.

## Why use Track B (vs the default Track A)

| Question | Track A (default) | Track B (BYOA harness) |
|---|---|---|
| What does `Module_MRP.bas` DO? | "data-loader, 80 lines. Reads sheet `订单`. Calls `Helper.Lookup`. Has `On Error Resume Next` at line 23." | "Implements classic MRP netting: nets gross requirements from MPS against on-hand inventory and open POs, expanded through BOM. Hardcodes a 0.05 loss factor." |
| What does the `MPS` sheet represent? | "Aggregator/output sheet — many of its formulas pull from other sheets." | "Master Production Schedule — the central planning grid that nets forecast demand against constraints to produce a weekly build plan." |
| Hardcode risk in the workbook? | "`_constants!C4` value `0.05` has fan-in 80; modifying it cascades to 80 formulas across MPS / 产能规划." | "The 0.05 LOSS_RATE in `_constants!C4` is the plant yield assumption. Single keystroke shifts every planned quantity by 5%." |

Track A is enough to identify pillar cells and fragile patterns. Track B
adds the **business-semantic layer** an outside reviewer or executive
needs to understand WHY the audit's findings matter.

## Compliance note (read this if your IT will ask)

`tier1-audit` makes **zero LLM API calls**. `--harness` produces text files;
`--ingest` parses text files. Both phases are pure local I/O. There is no
`requests` / `httpx` / `anthropic` / `openai` import anywhere in the
package — see `tests/test_corpus.py::test_zero_llm_in_source` and
`tests/test_harness.py::test_harness_zero_llm_imports` for guarded checks.

The LLM call happens **inside YOUR existing tooling** (VS Code Copilot Chat,
Claude Desktop, ChatGPT web). You paste; the LLM responds; you save.
Your IT department's compliance review for the audit tool itself sees only
file I/O.

If your company already approves Microsoft Copilot or your enterprise
ChatGPT account, BYOA reuses that approval — you don't need a second
vendor review for an LLM proxy that doesn't exist.

## Prerequisites

Pick whichever you prefer (any chat-style LLM client works):

- **VS Code Copilot Chat** (recommended for most users) — install the
  GitHub Copilot extension, sign in, ensure you can switch the model to
  Claude Sonnet 4.6 or GPT-4 family.
- **Claude Desktop** (free download from anthropic.com) — drag-and-drop
  attachments are easiest here.
- **`claude` CLI** (`npm i -g @anthropic-ai/claude-code` or similar) —
  for shell-only workflows.
- **ChatGPT web** (chatgpt.com) — works fine for non-confidential workbooks;
  attach `dossier.json` and paste `prompt.md`.

You also need:

- Python 3.9+ with `tier1-audit` installed (`pip install tier1-audit`)
- Your `.xlsm` workbook accessible on disk
- A working text editor (any) to save the LLM's JSON response

## Step 1 — Run extract

```bash
audit-xlsm path/to/your-file.xlsm --harness --out-dir ./audit-output
```

Expected output:

```
input    : your-file.xlsm
out-dir  : /abs/path/to/audit-output
format   : both
sanitize : off
complex. : 87/100
smells   : 154
pillars  : 2
anomalies: 0
vba clf  : 40 module(s) classified
domain   : capacity-planning (confidence: medium, hits: 4)
sheets   : 12
formulas : 4585
vba      : 40 modules / 14709 lines
errors   : 0
harness  : dossier.json + prompt.md (25 marker(s) to fill)
```

Files written to `./audit-output/`:

- `audit.md`, `audit.html`, `audit.json` — the standard Track A audit
- `dossier.json` — structured workbook context (sheets, full VBA source,
  pillars, smells, magic numbers, marker list with per-marker context)
- `prompt.md` — the mega-prompt you paste into the LLM

The line `harness  : dossier.json + prompt.md (25 marker(s) to fill)`
tells you how many semantic narratives the LLM will be asked to write.

### Privacy: --sanitize before --harness

If your workbook contains sensitive cell values, use `--sanitize` to
redact every non-formula cell before the dossier is built. Formulas, VBA
source, structural counts, and pillar metadata are preserved (those are
what the LLM needs):

```bash
audit-xlsm path/to/sensitive.xlsm --harness --sanitize --out-dir ./audit-output
```

The dossier you paste into your LLM will then contain `<redacted>` instead
of cell values. Same audit quality minus the data leak risk.

## Step 2 — Open Copilot Chat (or your LLM of choice)

### VS Code Copilot Chat

1. Open VS Code in any folder (you don't need to open the audit folder).
2. Press **Cmd+Alt+I** (macOS) or **Ctrl+Alt+I** (Windows/Linux) to open
   Copilot Chat.
3. In the chat input area, click the model picker (lower-right of the
   chat panel). Pick **Claude Sonnet 4.6** if available, otherwise the
   highest-capability model your subscription includes.

### Claude Desktop

1. Open Claude Desktop.
2. Start a new conversation.
3. (Optional) verify you're on the latest Claude Sonnet model.

### ChatGPT web

1. Open chatgpt.com, start a new chat.
2. Pick GPT-4 / o1 / whichever model your subscription includes. Avoid the
   "GPT-3.5" tier — narrative quality drops noticeably.

## Step 3 — Paste the prompt + attach the dossier

Here's the universal flow that works across all clients:

1. **Open `audit-output/prompt.md` in any text editor.**
   - macOS: `open audit-output/prompt.md`
   - Windows: `start audit-output\prompt.md`
   - or just open it from the editor's file picker.
2. **Select all (Cmd+A / Ctrl+A) and copy (Cmd+C / Ctrl+C).**
3. **Paste into the chat input.**
4. **Attach `audit-output/dossier.json`**:
   - VS Code Copilot Chat: drag the file from VS Code's Explorer pane into
     the chat input. (Alternative: click the paperclip icon if available.)
   - Claude Desktop: drag the file from Finder/Explorer into the chat
     window. (Alternative: paste the file's CONTENT below the prompt
     wrapped in a code fence — `dossier.json` content can also be inlined.)
   - ChatGPT: click the paperclip / attach button, select `dossier.json`.
5. **Send.**

The prompt itself reminds the LLM that `dossier.json` is attached and
restates each question's context inline, so the LLM can answer even if
attachments aren't visible in your client.

## Step 4 — Get the JSON response back

The LLM should respond with **one JSON object** like:

```json
{
  "data-flow:MPS": "Master Production Schedule — the central planning grid that nets forecast demand against constraints. Inputs: forecast loss factor from `_constants!C4`, demand from `订单`. Feeds 产能规划 and MRP downstream.",
  "vba-narration:Module_Main": "Implements classic MRP netting: gross requirements from MPS minus on-hand inventory and open POs, expanded through BOM. Hardcodes 0.05 loss factor.",
  "domain-method:manufacturing-capacity-planning": "Strong evidence this is a manufacturing capacity-planning workbook: BOM + MPS + MRP sheet trio is the textbook decomposition. Highest-priority hardcode risk: the 0.05 LOSS_RATE in `_constants!C4`.",
  ...
}
```

What to verify in the response:
- It's a top-level `{...}` object, NOT an array, NOT wrapped in extra
  metadata.
- Keys are marker IDs from `prompt.md`'s "Questions" section.
- Values are short narrative strings (each ~2-4 sentences).
- No markdown fences. (If your LLM wraps the JSON in <code>```json ... ```</code>,
  the ingest step will tolerate that — but cleaner JSON is better.)

If the LLM only answered some markers, that's fine. The ingest step
**gracefully degrades**: any marker without a response keeps its heuristic
narrative.

## Step 5 — Save the response as `responses.json`

Copy the LLM's full JSON output (just the JSON, not surrounding chatter)
and save it as `audit-output/responses.json`:

- macOS / Linux: `pbpaste > audit-output/responses.json` (after copying)
- Windows: open Notepad, paste, save as `responses.json` (Save as type:
  All Files; encoding: UTF-8).
- VS Code: New File → paste → Save As `responses.json`.

Sanity check the file:

```bash
python -c "import json; print(len(json.load(open('audit-output/responses.json'))), 'narratives')"
```

If that prints a number, you have valid JSON. If it errors, the LLM
likely included markdown fences or extra text. Trim to just the `{...}`
object and re-save.

## Step 6 — Run ingest

```bash
audit-xlsm path/to/your-file.xlsm --ingest audit-output/responses.json --out-dir audit-output
```

Expected output (yours will differ in counts):

```
input        : your-file.xlsm
out-dir      : /abs/path/to/audit-output
responses    : audit-output/responses.json
enriched md  : /abs/path/to/audit-output/audit-enriched.md
enriched html: /abs/path/to/audit-output/audit-enriched.html
replaced     : 17 marker(s)
kept (heur.) : 8 marker(s) with no LLM response — heuristic narrative kept
```

The ingest is purely textual: it finds each `<!-- LLM-AUGMENT: ID -->`
comment in the original `audit.md` / `audit.html` and replaces the
heuristic prose immediately after with the LLM's narrative. The marker
itself is preserved (so re-running ingest is idempotent).

## Step 7 — Open the enriched audit

```bash
open audit-output/audit-enriched.html
```

Compare side-by-side with the original `audit.html` — you'll see the
heuristic narratives ("data-loader, 80 lines, reads sheet X") replaced
with semantic ones ("Implements classic MRP netting, hardcodes 0.05 loss
factor"). The structural skeleton (sheet headings, module headings,
tables, diagrams) is unchanged — only the prose paragraphs differ.

## Troubleshooting

### "FATAL: responses file is not valid JSON"

The most common cause: your LLM wrapped the JSON in markdown fences:

````
```json
{ "data-flow:MPS": "..." }
```
````

The ingest tolerates this single common case (it strips fence markers
before parsing). If it still fails, open `responses.json` and verify:
- The first non-whitespace character is `{`.
- The last non-whitespace character is `}`.
- No prose before/after the JSON object.

If you're sure the JSON is valid, paste it into a JSON validator
([jsonlint.com](https://jsonlint.com), or `python -c "import json; json.load(open('responses.json'))"`)
to identify the exact syntax issue. The error message includes line
+ column from `json.JSONDecodeError`.

### "INFO: marker 'X' had no LLM response — keeping heuristic narrative"

This is **graceful degradation** working as designed. Some markers in your
audit didn't get a corresponding key in `responses.json`. You have three
options:

1. **Ignore it.** The original heuristic narrative stays in the enriched
   output. Partial enrichment is still progress.
2. **Re-run the LLM** with a prompt asking specifically for the missing
   markers. You don't need to redo the whole prompt — just paste a
   shortened version listing the missing IDs and ask for those narratives.
3. **Save a NEW responses.json with the additions** and re-run ingest —
   it's idempotent, so this is safe.

### "INFO: ignored unused response key 'X' (not present in audit.md)"

The LLM hallucinated a marker ID that doesn't exist in your audit.
Harmless — just ignored. Common when the LLM extrapolates from your
example narratives.

### Markers not being substituted (no replacement happens)

Double-check that the `audit.md` you're ingesting against is the same one
the marker IDs came from. If you regenerated the audit with different
options (e.g. switched between sanitize on/off, or added/removed
sheets), the marker IDs may have shifted. Re-run `--harness` to get a
fresh `dossier.json` + `prompt.md` and use the matching `audit.md`.

### LLM refuses to write narratives ("I don't see attachments")

Some LLM clients don't pass attachments through to the model reliably.
The prompt already restates each question's context inline, so attachment
visibility is nice-to-have, not required. If the LLM insists, you can
also paste `dossier.json` content directly below the prompt wrapped in
a code fence — but watch the prompt size cap (some clients reject very
long prompts).

### Output narratives are mechanical, not semantic

If the LLM is producing "loops over rows 2 to 200 and writes column G",
the rules section of the prompt isn't getting through. Try:

- Switch to a higher-capability model (Sonnet 4.x / GPT-4 / o1 family).
- Re-emphasize: "Read the Rules section again. Each narrative MUST be
  business-semantic, not mechanical. If you can't identify a method,
  say so — but don't describe loops."
- Re-paste the prompt; some long-context models drop earlier instructions.

## Re-running ingest after editing responses.json

Ingest is idempotent: running it twice on the same `responses.json`
produces a byte-identical `audit-enriched.md`. So you can iterate:

1. Run ingest, open the enriched output.
2. Notice a narrative is wrong / missing / could be better.
3. Re-prompt the LLM for just that ID.
4. Update `responses.json` with the new narrative.
5. Re-run `audit-xlsm <file> --ingest responses.json --out-dir audit-output`.

The marker comments stay in the enriched output, so subsequent ingests
re-find them and re-substitute.
