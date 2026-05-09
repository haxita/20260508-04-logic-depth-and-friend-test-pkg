# Agent-harness research — VS Code Copilot 2026 + audit-xlsm Track B

Date: 2026-05-09  ·  Time-boxed at ~45 min wall-clock

## Q1.1 — VS Code Copilot in 2026: what auto-loads from a repo

**Custom agents (formerly chatmodes) are the right hook.** As of VS Code
1.1xx (April 2026), `.chatmode.md` was renamed to `.agent.md`; legacy
files are migrated by rename ([custom-agents][vscode-agents]).

Auto-loaded from repo root, in priority order:

1. `.github/copilot-instructions.md` — applied to every chat request
   in the workspace ([instructions][copilot-inst]).
2. `.github/instructions/*.instructions.md` — scoped via `applyTo:`.
3. `.github/agents/<name>.agent.md` — custom agents, auto-discovered.
4. `AGENTS.md` / `CLAUDE.md` at root — also recognised.
5. `.claude/agents/<name>.md` — Claude Code agents (read by Copilot too).
6. `.vscode/mcp.json` — local MCP server registry ([mcp][mcp-vsc]).

YAML frontmatter for `.agent.md` (canonical [GitHub spec][gh-spec]):
`description` (required), `name`, `tools`, `model`, `target` (`vscode`
or `github-copilot`), `mcp-servers`, `handoffs`. The `tools:` array
accepts aliases: `read`, `edit`, `search`, `execute`, `web`, `todo`,
`agent`. **`execute` runs shell commands** in the user's default
terminal — that's our key affordance.

Invocation: agent dropdown in Copilot Chat, `@<agent-name>` mention,
or `/agent` slash in Copilot CLI ([CLI custom agents][cli-agents]).
"One prompt → multi-step tool execution → final output" is the headline
use case. Reliability caveat: outputs are non-deterministic; the agent
may need a re-prompt mid-flow if it loses context.

## Q1.2 — Claude Code surfaces

Three under `.claude/`: `agents/<name>.md` ([subagents][cc-subagents]),
`skills/<name>/SKILL.md` (auto-trigger is **acknowledged flaky** in 2026
community ([trigger issues][skill-trigger])), and `commands/<name>.md`
(legacy but still working). Slash commands' `!`-prefix runs a shell
command and inlines stdout into the prompt ([slash][cc-slash]) — the
most reliable Claude Code surface for our case.

## Q1.3 — MCP servers

VS Code reads `.vscode/mcp.json` (`servers` key, NOT `mcpServers`).
Install friction for a local stdio server: edit `mcp.json`, trust
dialog, possible `Developer: Reload Window`, agent mode required
([VS Code MCP][mcp-vsc]). Precedents for wrapping a Python CLI as an
MCP server: [`mcp-tools-py`][mcp-tools-py] (pylint/pytest/mypy),
[`code-analysis-mcp`][code-analysis-mcp] (repo-scoped analysis),
[`codepathfinder`][codepathfinder] (codebase Q&A). Feasible but adds a
second install + config edit + protocol versioning. **Not the lightest
path for first contact.**

## Q1.4 — Real precedents

- [`harness/harness-skills`][harness-skills]: `.github/copilot-instructions.md`
  + skill collection drives CI/CD ops from a single chat prompt.
- [`github/awesome-copilot`][awesome-copilot]: community library of
  agents, skills, instructions; `agents.instructions.md` is the style
  guide.
- [`githubnext/agentics`][agentics]: 30+ agentic workflows (CI Doctor,
  Weekly Repo Map, Link Checker) — single-prompt → shell-runs →
  markdown-report pattern identical to ours.

## Q1.5 — Recommendation for OUR use case

**Hybrid: ship a `.github/agents/xlsm-audit.agent.md` custom agent + a
project-level `.github/copilot-instructions.md` + a parallel
`.claude/agents/xlsm-auditor.md` for the rare Claude-Code user, AND keep
the manual `--harness`/`--ingest` flow as the documented fallback.**

Why this combination, not pure MCP nor pure CLI:

1. **Friend's stated environment is VS Code + Copilot Chat.** Custom
   agents are the auto-discovered, zero-config surface — the friend opens
   the repo, clicks the agent dropdown, picks "xlsm-audit", types a path,
   waits. No `mcp.json`, no second install, no protocol.
2. **`tools: ['execute', 'read', 'edit']`** gives the agent shell + file
   I/O, which is everything we need: run `audit-xlsm --harness`, read
   `dossier.json` + `prompt.md`, apply the prompt itself (Copilot is the
   LLM), write `responses.json`, run `audit-xlsm --ingest`, open the HTML.
   No tool itself ever calls an LLM; we keep BYOA architecturally.
3. **Reliability is non-zero but not bulletproof.** Agent mode is
   non-deterministic ([VS Code agents overview][vscode-agents]); some
   friends will hit a step where Copilot returns prose instead of JSON,
   or skips the ingest call. **Therefore the manual `--harness`/`--ingest`
   path stays — it's the always-works floor.**
4. **MCP is over-engineered for v1.** It gains us nothing the agent file
   doesn't already provide and adds a second install step, a second trust
   dialog, and a third surface to teach. Defer to v2 if friends actually
   ask "can I get tools out of this from non-Copilot clients?"

Concretely: Phase 2 ships three config files plus updated docs. The
friend's path becomes:

> 1. `pip install .`
> 2. Open repo in VS Code.
> 3. Open Copilot Chat → agent dropdown → **xlsm-audit**.
> 4. Type the file path.
> 5. Wait. Open `audit-output/audit-enriched.html`.

If the agent flow stalls, the friend falls back to `audit-xlsm --harness`
+ paste-into-Copilot Chat + `audit-xlsm --ingest`, which is what the v0.1
docs already document.

[vscode-agents]: https://code.visualstudio.com/docs/copilot/customization/custom-agents
[vscode-chatmodes]: https://code.visualstudio.com/docs/copilot/customization/custom-chat-modes
[copilot-inst]: https://code.visualstudio.com/docs/copilot/customization/custom-instructions
[mcp-vsc]: https://code.visualstudio.com/docs/copilot/customization/mcp-servers
[gh-spec]: https://docs.github.com/en/copilot/reference/custom-agents-configuration
[cli-agents]: https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/create-custom-agents-for-cli
[cc-subagents]: https://code.claude.com/docs/en/sub-agents
[cc-slash]: https://code.claude.com/docs/en/slash-commands
[skill-trigger]: https://dev.to/oluwawunmiadesewa/claude-code-skills-not-triggering-2-fixes-for-100-activation-3b57
[mcp-tools-py]: https://github.com/marcusjellinghaus/mcp-tools-py
[code-analysis-mcp]: https://github.com/saiprashanths/code-analysis-mcp
[codepathfinder]: https://codepathfinder.dev/mcp
[harness-skills]: https://github.com/harness/harness-skills
[awesome-copilot]: https://github.com/github/awesome-copilot
[agentics]: https://github.com/githubnext/agentics
