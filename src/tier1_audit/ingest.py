"""Track B (BYOA) ingest phase — substitute LLM responses into the audit.

Architectural promise: this module never calls an LLM, never opens a network
socket, and never imports any LLM SDK. It only reads a JSON file the user
manually saved (after pasting the prompt into their own LLM client) and
substitutes its values into audit.md / audit.html, producing
audit-enriched.md / audit-enriched.html.

Substitution semantics:
    - The marker `<!-- LLM-AUGMENT: ID -->` is the anchor.
    - Everything between the marker and the next "section break" is the
      heuristic narrative we replace.
    - Section break = the next line that is one of:
        * another `<!-- LLM-AUGMENT: ... -->`
        * a markdown heading at level 2 or 3 (`## ` or `### `)
        * a horizontal rule (`---`) — used between major sections sometimes
        * end of file
    - The marker line itself is PRESERVED (so re-ingest is idempotent).

Graceful degradation:
    - If a marker has no entry in responses.json: keep heuristic narrative,
      log to stderr.
    - If responses.json has IDs not present in audit.md: ignore + log.
    - If responses.json is malformed JSON: raise SystemExit with a clear
      message including line number.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

_RE_MD_MARKER_LINE = re.compile(r"^\s*<!--\s*LLM-AUGMENT:\s*([^\s][^>]*?)\s*-->\s*$")
# Loose marker (anywhere on the line) — used in HTML where comments may
# coexist with adjacent tags on the same line in some renderings.
_RE_MD_MARKER_ANY = re.compile(r"<!--\s*LLM-AUGMENT:\s*([^\s][^>]*?)\s*-->")


# =============================================================================
# Loading + validating the user's JSON
# =============================================================================

def load_responses(path) -> dict:
    """Read responses.json and return {id: narrative}. Validates structure.

    Raises SystemExit with a helpful message on malformed JSON. Strips
    common pitfalls: leading/trailing whitespace per value, empty narratives,
    non-string values.
    """
    p = Path(path)
    if not p.exists():
        raise SystemExit(f"FATAL: responses file not found: {p}")
    raw = p.read_text(encoding="utf-8")
    # Tolerate the LLM accidentally wrapping output in ```json ... ``` fences.
    stripped = raw.strip()
    if stripped.startswith("```"):
        # Strip first fence line and last fence line
        lines = stripped.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        # Drop trailing fence
        while lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        stripped = "\n".join(lines).strip()
    try:
        data = json.loads(stripped)
    except json.JSONDecodeError as e:
        raise SystemExit(
            f"FATAL: responses file is not valid JSON: {p}\n"
            f"  json.JSONDecodeError: {e.msg}\n"
            f"  at line {e.lineno} column {e.colno} (char {e.pos})\n"
            f"  Tip: the LLM may have added a markdown fence (```json ... ```) — "
            f"strip it and retry, or paste only the JSON object."
        )
    if not isinstance(data, dict):
        raise SystemExit(
            f"FATAL: responses file must be a JSON OBJECT (dict), got "
            f"{type(data).__name__}: {p}"
        )
    # Coerce values to stripped strings; drop empties/non-strings (logged).
    cleaned: dict = {}
    skipped: list = []
    for k, v in data.items():
        if not isinstance(k, str):
            skipped.append(f"non-string key {k!r}")
            continue
        if not isinstance(v, str):
            skipped.append(f"key {k!r}: value type {type(v).__name__} (expected string)")
            continue
        s = v.strip()
        if not s:
            skipped.append(f"key {k!r}: empty narrative")
            continue
        cleaned[k] = s
    if skipped:
        for msg in skipped:
            print(f"WARN: responses.json: skipped {msg}", file=sys.stderr)
    return cleaned


# =============================================================================
# Markdown substitution
# =============================================================================

def _is_section_break_md(line: str) -> bool:
    """True if this line ends a heuristic-narrative block in markdown.

    Section break = next LLM-AUGMENT marker, next H2 (`## `), or `---` rule.
    H3 (`### `) is NOT a section break: in our renderer, an H3 heading
    immediately follows a marker as the structural identifier of the block
    (e.g. `### \`MPS\` (visible, ...)`) — we preserve that heading and
    only replace the prose that follows it.
    """
    s = line.lstrip()
    if _RE_MD_MARKER_LINE.match(line):
        return True
    if s.startswith("## ") and not s.startswith("### "):
        return True
    if s.startswith("---"):
        return True
    return False


def substitute_markers(audit_md_text: str, responses: dict) -> tuple:
    """Replace heuristic narratives following each marker with LLM responses.

    Returns (new_text, stats) where stats = {
        'replaced': [marker_ids],
        'kept_heuristic': [marker_ids missing from responses],
        'unused_responses': [response_ids not in audit_md],
    }

    The marker line itself is preserved (so re-ingest is idempotent). The
    H3 heading that immediately follows the marker (if any) is also preserved
    — it carries the structural identity (sheet name, module name) that must
    survive substitution. The prose paragraphs / bullet content AFTER the
    optional heading are replaced with the LLM response.
    """
    lines = audit_md_text.splitlines()
    out: list = []
    n = len(lines)
    i = 0

    seen_markers: set = set()
    replaced: list = []
    kept_heuristic: list = []

    while i < n:
        line = lines[i]
        m = _RE_MD_MARKER_LINE.match(line)
        if not m:
            out.append(line)
            i += 1
            continue

        marker_id = m.group(1).strip()
        seen_markers.add(marker_id)
        # Always keep the marker line itself
        out.append(line)
        i += 1

        # Optional: if the next line is an H3 heading, keep it too — it's
        # the structural identifier for this section, not part of the
        # narrative.
        if i < n and lines[i].lstrip().startswith("### "):
            out.append(lines[i])
            i += 1

        # Find the end of the heuristic narrative block: until next
        # marker / H2 / horizontal rule.
        block_start = i
        while i < n and not _is_section_break_md(lines[i]):
            i += 1
        block_end = i  # exclusive

        if marker_id in responses:
            # Replace the heuristic block with the LLM narrative.
            # Preserve trailing blank line if the heuristic had one (so the
            # rendered markdown spacing stays consistent).
            had_trailing_blank = (
                block_end > block_start and lines[block_end - 1].strip() == ""
            )
            out.append(responses[marker_id])
            if had_trailing_blank:
                out.append("")
            replaced.append(marker_id)
        else:
            # Keep heuristic
            for j in range(block_start, block_end):
                out.append(lines[j])
            kept_heuristic.append(marker_id)

    # Identify unused responses
    unused = sorted(set(responses.keys()) - seen_markers)

    new_text = "\n".join(out)
    # Preserve trailing newline behavior of the input
    if audit_md_text.endswith("\n") and not new_text.endswith("\n"):
        new_text = new_text + "\n"
    elif not audit_md_text.endswith("\n") and new_text.endswith("\n"):
        new_text = new_text.rstrip("\n")

    stats = {
        "replaced": sorted(replaced),
        "kept_heuristic": sorted(kept_heuristic),
        "unused_responses": unused,
    }
    return new_text, stats


# =============================================================================
# HTML substitution
# =============================================================================

# In the HTML output, the renderer passes <!-- LLM-AUGMENT: ID --> through
# as a literal comment line (see render._md_lines_to_html). The narrative
# that follows is rendered as one or more <p>, <ul>, or <h3> blocks, until
# the next marker comment or </section> / next <h2>/<h3>.
#
# Substitution strategy: walk lines, detect marker lines, then consume HTML
# blocks until we hit a section-break HTML feature.

# Patterns for the HTML substitution. Unlike markdown, HTML emitted by the
# renderer is mostly on a single long line, so we operate on the full text
# with regex rather than line-by-line.
_RE_HTML_MARKER = re.compile(
    r"<!--\s*LLM-AUGMENT:\s*([^\s][^>]*?)\s*-->", re.DOTALL,
)
# A "section break" in HTML means: another marker, an H2 opening tag, an
# <hr>, or </section>. We DON'T break on H3 — the renderer emits H3 as the
# structural identifier of the marker block and we want to keep it.
_RE_HTML_BREAK = re.compile(
    r"(<!--\s*LLM-AUGMENT:[^>]+?-->|<h2\b|<hr\b|</?section\b)",
    re.IGNORECASE | re.DOTALL,
)


def _escape_html(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def substitute_markers_html(audit_html_text: str, responses: dict) -> tuple:
    """Replace heuristic narratives in HTML output (regex-based, one-pass).

    The renderer emits marker comments inline within longer lines. We walk
    the text, find each marker, optionally skip an immediately-following
    `<h3>...</h3>` (structural identifier), then replace the heuristic
    block up to the next break with `<p class="llm-narrative">...`.
    """
    seen_markers: set = set()
    replaced: list = []
    kept_heuristic: list = []

    out_parts: list = []
    pos = 0

    for m in _RE_HTML_MARKER.finditer(audit_html_text):
        marker_id = m.group(1).strip()
        seen_markers.add(marker_id)

        # Append everything before this marker
        out_parts.append(audit_html_text[pos:m.start()])
        # Append the marker comment itself (preserved for re-ingest)
        out_parts.append(m.group(0))
        cursor = m.end()

        # Optional: skip and preserve a leading <h3 ...>...</h3>
        h3_match = re.match(
            r"\s*<h3\b[^>]*>.*?</h3>",
            audit_html_text[cursor:],
            re.IGNORECASE | re.DOTALL,
        )
        if h3_match:
            out_parts.append(audit_html_text[cursor:cursor + h3_match.end()])
            cursor += h3_match.end()

        # Find the next break starting at `cursor`
        nxt = _RE_HTML_BREAK.search(audit_html_text, cursor)
        block_end = nxt.start() if nxt else len(audit_html_text)

        if marker_id in responses:
            narrative = responses[marker_id]
            out_parts.append(
                f'<p class="llm-narrative">{_escape_html(narrative)}</p>'
            )
            replaced.append(marker_id)
        else:
            # Keep the heuristic block verbatim
            out_parts.append(audit_html_text[cursor:block_end])
            kept_heuristic.append(marker_id)

        pos = block_end

    out_parts.append(audit_html_text[pos:])
    new_text = "".join(out_parts)

    unused = sorted(set(responses.keys()) - seen_markers)

    stats = {
        "replaced": sorted(replaced),
        "kept_heuristic": sorted(kept_heuristic),
        "unused_responses": unused,
    }
    return new_text, stats


# =============================================================================
# Public entry point
# =============================================================================

def ingest(audit_md_path, audit_html_path, responses_path,
           out_md_path, out_html_path) -> dict:
    """Ingest LLM responses into audit.md / audit.html.

    Args are paths (str or Path). Returns a stats dict combining markdown +
    html substitution outcomes.

    audit_html_path / out_html_path may be None — in which case only md
    is processed.
    """
    md_path = Path(audit_md_path)
    if not md_path.exists():
        raise SystemExit(f"FATAL: audit.md not found: {md_path}")
    out_md = Path(out_md_path)

    responses = load_responses(responses_path)

    md_text = md_path.read_text(encoding="utf-8")
    new_md, md_stats = substitute_markers(md_text, responses)
    out_md.parent.mkdir(parents=True, exist_ok=True)
    out_md.write_text(new_md, encoding="utf-8")

    # Log graceful-degradation events
    for mid in md_stats["kept_heuristic"]:
        print(f"INFO: marker '{mid}' had no LLM response — keeping heuristic narrative",
              file=sys.stderr)
    for mid in md_stats["unused_responses"]:
        print(f"INFO: ignored unused response key '{mid}' (not present in audit.md)",
              file=sys.stderr)

    html_stats = None
    if audit_html_path is not None and out_html_path is not None:
        html_path = Path(audit_html_path)
        if html_path.exists():
            html_text = html_path.read_text(encoding="utf-8")
            new_html, html_stats = substitute_markers_html(html_text, responses)
            out_html = Path(out_html_path)
            out_html.parent.mkdir(parents=True, exist_ok=True)
            out_html.write_text(new_html, encoding="utf-8")

    return {
        "md_path": str(out_md),
        "html_path": str(out_html_path) if out_html_path else None,
        "md_replaced": md_stats["replaced"],
        "md_kept_heuristic": md_stats["kept_heuristic"],
        "md_unused_responses": md_stats["unused_responses"],
        "html_replaced": (html_stats or {}).get("replaced"),
    }
