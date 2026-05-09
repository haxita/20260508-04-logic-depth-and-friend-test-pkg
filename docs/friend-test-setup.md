# Friend-test setup

Hi! Thanks for trying `tier1-audit`. This document walks you through getting
it running on your work laptop, then running it against one of your real
`.xlsm` files. The tool runs **entirely on your machine** — there is no cloud
component, no telemetry, and no LLM API call.

## What you'll need

- **Python 3.9 or newer** (the Python that ships with VS Code's Python
  extension is fine; or any system Python; macOS's bundled `python3` works).
  Verify with `python --version` (or `python3 --version` on macOS / Linux).
- **Git** (only if you choose Path 1 with `git clone`, or Path 2 from GitHub).
- **VS Code** is recommended but not required. Any terminal works.

If you only have Python 3.8 or older, ask IT for a 3.9+ install — every
maintained Python install in the last 4 years works.

---

## Install — pick the path that matches your environment

There are four install paths below. **Most friends should use Path 1.**
Only switch paths if Path 1's error message points you somewhere else.

### Path 1 (default, works with old pip) — `pip install .`

This is the simplest path. Works with the legacy pip that ships on macOS
system Python 3.9 and many corporate Windows images. **No editable install
needed**, so you avoid the PEP 660 / setuptools-backend issues that often
trip old pip.

#### Option A — git clone

```bash
git clone https://github.com/haxita/xlsm-audit.git
cd xlsm-audit
python -m venv .venv
# macOS/Linux:
source .venv/bin/activate
# Windows PowerShell:
# .venv\Scripts\Activate.ps1
# Windows cmd:
# .venv\Scripts\activate.bat
pip install .
audit-xlsm path/to/your-file.xlsm
```

#### Option B — zip download (no Git required)

If you don't have Git installed, download the repo zip from GitHub
("Code" -> "Download ZIP"), unzip it, and:

```bash
cd <unzipped-folder>
python -m venv .venv
source .venv/bin/activate          # or the Windows variants above
pip install .
audit-xlsm path/to/your-file.xlsm
```

### Path 2 (online direct, modern pip) — `pip install git+https://...`

If you have a recent pip (>= 21.3) and direct internet access, you can
skip the clone step and let pip do everything:

```bash
pip install git+https://github.com/haxita/xlsm-audit.git
audit-xlsm path/to/your-file.xlsm
```

After this the `audit-xlsm` command is on your `PATH`.

### Path 3 (sealed network — internal Artifactory / Nexus / corporate mirror)

Inside companies with no public PyPI access, you typically have an
internal package mirror (Artifactory, Nexus, devpi, etc.). Point pip at
that mirror:

```bash
pip install --index-url https://<INTERNAL_MIRROR_URL>/simple/ .
```

Or for the GitHub-direct flavour, ask IT to mirror the repo into your
internal Git server, then:

```bash
pip install --index-url https://<INTERNAL_MIRROR_URL>/simple/ \
    git+https://<INTERNAL_GIT>/path/to/tier1-audit.git
```

The runtime dependencies that must be available on your internal mirror:

- `openpyxl` (>= 3.1)
- `oletools` (>= 0.60)
- `formulas` (>= 1.3)

These are common — most Artifactory / Nexus instances already mirror them.
If they're missing, ask IT to add them; they are all permissively-licensed
packages used widely in industry.

### Path 4 (corporate proxy)

If your laptop reaches the public internet only through a corporate HTTPS
proxy:

```bash
pip install --proxy http://corp.proxy:port .
```

Or with environment variables:

```bash
export HTTPS_PROXY=http://corp.proxy:port
export HTTP_PROXY=http://corp.proxy:port
pip install .
```

If your proxy uses authentication, the form is
`http://user:pass@corp.proxy:port`. Most companies provide an internal
documentation page for the exact value.

---

## Troubleshooting

| Error you saw | Most likely cause | What to switch to |
|---|---|---|
| `error: Project file ... has 'requires-python' that doesn't match installed Python` | Python is too old (< 3.9) | Install Python 3.9+; ask IT |
| `ERROR: Project file has a 'pyproject.toml' and its build backend is missing the 'build_editable' hook` | Old pip without PEP 660 editable support | Path 1: use `pip install .` (not `-e .`) |
| `WARNING: Retrying (Retry(total=4, ...)) after connection broken by 'NewConnectionError'` | No public-internet access | Path 3 (internal mirror) or Path 4 (proxy) |
| `Could not find a version that satisfies the requirement openpyxl>=3.1` | Internal mirror missing the package | Ask IT to add `openpyxl`, `oletools`, `formulas` to the mirror |
| `ssl.SSLCertVerificationError: certificate verify failed` | Corporate MITM proxy with self-signed cert | Path 4 with `--trusted-host` flag, or have IT add the cert to the system trust store |
| `running scripts is disabled on this system` | Windows PowerShell ExecutionPolicy | See "Windows-specific notes" below |
| `RuntimeError: Failed to extract VBA macros` | File is encrypted / password-protected | Save-as without password, or skip the file |
| `ImportError: cannot import name 'X' from 'numpy'` | Pre-installed numpy mismatch with `formulas` | `pip install --upgrade --force-reinstall numpy formulas` |

---

## First run — privacy mode recommended

For your first run on a real file, **use sanitize mode**. It produces the
exact same report shape but redacts every non-formula cell value:

```bash
audit-xlsm path/to/your-file.xlsm --sanitize
```

The output `out/audit.md` will start with:

> 🔒 SANITIZED MODE — no cell values in this report.

This lets you read the audit, share excerpts with Michael, or paste sections
into a GitHub issue without worrying about leaking proprietary data. The
formulas, VBA source, smells, structural counts are all preserved — that's
where the diagnostic value is.

After you've read the sanitized version and confirmed it's safe, you can
re-run without `--sanitize` for a richer report (still on your machine).

## Output formats

By default, the tool writes both `audit.md` and `audit.html` (and `audit.json`).

```bash
# md + html (default)
audit-xlsm file.xlsm

# md only
audit-xlsm file.xlsm --format md

# html only
audit-xlsm file.xlsm --format html
```

The HTML version embeds Mermaid diagrams via the public CDN so you can read
the report in any browser and print to PDF (browser ⌘+P → Save as PDF). If
your machine has no internet access at view time:

```bash
audit-xlsm file.xlsm --mermaid-inline
```

This downloads `mermaid.min.js` once (during the install of dependencies)
and inlines it into the HTML, so the file is fully self-contained.

## Compliance reasoning (for your IT team if asked)

If your IT or security team needs justification for installing this tool,
the relevant facts:

- **All Python source is open and inspectable.** The package is < 2,500
  lines of Python in `src/tier1_audit/`. Anyone can read it.
- **Three open-source dependencies**, all widely-deployed:
  - `openpyxl` (MIT) — reads `.xlsm` files. Used by Microsoft's own
    documentation as the recommended Python interop library. ~10M
    downloads/month.
  - `oletools` (BSD) — extracts VBA from OLE containers. Used by SOC
    teams and malware researchers; vetted security tooling.
  - `formulas` (EUPL-1.1) — Excel formula tokenizer. **We use the
    tokenizer only.** No formula is ever evaluated.
- **No network calls during audit.** Verifiable in two ways:
  1. Run with the network disabled — the audit completes normally.
  2. `pip install` itself reaches out to PyPI (or your internal mirror)
     once to fetch the dependencies. After install, no further network
     access is needed or made by the audit pipeline.
- **No data sent to any third party.** All output is written to local
  disk, in the directory you pass to `--out-dir` (default: `./out/`).
- **No LLM / AI service.** This is pure static analysis — pattern
  matching, tokenization, counting. There are no API keys to provide.
- **No macro is executed.** VBA source is parsed as text, not run.
- **Runs on stock Python 3.9+.** No exotic runtime, no compiled
  binaries beyond the standard wheels for `cryptography` (a transitive
  dep of `oletools`).
- **`--sanitize` mode exists exactly for this scenario:** generate the
  report, redact cell values, share the redacted version with the
  vendor. We designed for the case where company files cannot leave the
  laptop.

## Windows-specific notes

### `ExecutionPolicy` (PowerShell)

If activating the venv triggers a "running scripts is disabled" error:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Then re-try `.venv\Scripts\Activate.ps1`. This change is for your user
only and is reversible.

### Paths with Chinese / Japanese / Korean characters or spaces

Quote the path:

```cmd
audit-xlsm "C:\Users\YourName\Documents\我的工作簿\report 2024.xlsm"
```

The tool is UTF-8 throughout — Chinese / Japanese / Korean sheet names and
formula contents are handled correctly.

### Long path limits

Windows historically capped paths at 260 characters. If your `.xlsm` lives
deep in a OneDrive sync folder and you hit a path-too-long error, copy the
file to a shorter location first (e.g. `C:\Temp\file.xlsm`).

## What I'm hoping you'll send back

After running the tool, what's most useful is your **qualitative reaction**
to the report. In particular:

- Did the **Pillar cells** section identify cells you yourself would have
  flagged as critical? Anything obvious it missed?
- Did the **Magic-number anomalies** section catch anything real, or was
  it noise? If both — the ratio matters.
- Are the **VBA module classification** labels accurate? "data-loader"
  and "transformer" are heuristic guesses; sometimes they're wrong, and
  knowing where they're wrong is gold.
- Did the **HTML report** print to a clean PDF? Did the Mermaid diagrams
  render correctly?
- How long did the audit take to run on your real file?
- Did anything crash, hang, or produce a clearly-wrong output?

You can:
- Paste a relevant excerpt from a `--sanitize`-mode `audit.md` into a
  GitHub issue, an email, or a chat message.
- Mention any sheet names / VBA module names that came out garbled (UTF-8
  edge cases).
- **Don't** send the original `.xlsm` file. We don't need it. The audit
  report is sufficient for diagnosis.

Thank you. This kind of feedback is exactly what shapes whether the next
version is genuinely useful to your team or just another tool that
half-works.
