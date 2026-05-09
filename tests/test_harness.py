"""Tests for the Track B (BYOA) harness — extract + ingest phases.

These tests guard the architectural promise: our tool never calls an LLM,
but PREPARES context the user can paste into theirs. The extract phase
emits dossier.json + prompt.md. The ingest phase substitutes the user's
saved JSON response back into audit.md / audit.html.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from tier1_audit.audit import build_audit
from tier1_audit.cli import main as cli_main
from tier1_audit.harness import (
    build_dossier,
    build_prompt,
    extract as harness_extract,
)
from tier1_audit.ingest import (
    ingest as harness_ingest,
    load_responses,
    substitute_markers,
)
from tier1_audit.render import render_markdown


REPO_ROOT = Path(__file__).resolve().parents[1]
CORPUS_BASE = REPO_ROOT.parent
SYNTH_PATH = CORPUS_BASE / "20260508-02-xlsb-extractor-mvp" / "test_files" / "capacity_planning_synth.xlsm"


def _require_synth():
    if not SYNTH_PATH.exists():
        pytest.skip(f"synth corpus not present at {SYNTH_PATH}")


def _run_harness_extract(tmp_path: Path) -> Path:
    """Helper: run the CLI with --harness and return the out_dir."""
    _require_synth()
    out = tmp_path / "out"
    out.mkdir(parents=True, exist_ok=True)
    rc = cli_main([str(SYNTH_PATH), "--out-dir", str(out), "--harness"])
    assert rc == 0
    return out


# ----------------------------------------------------------------------------
# 1. Extract creates dossier + prompt + standard audit outputs
# ----------------------------------------------------------------------------

def test_harness_extract_creates_dossier_and_prompt(tmp_path: Path):
    out = _run_harness_extract(tmp_path)
    # All five files written
    for fname in ("audit.md", "audit.html", "audit.json",
                  "dossier.json", "prompt.md"):
        assert (out / fname).exists(), f"missing {fname}"

    # dossier.json parses as JSON with the expected top-level keys
    dossier = json.loads((out / "dossier.json").read_text(encoding="utf-8"))
    expected_keys = {
        "schema_version", "workbook_meta", "sheets", "vba_modules",
        "workflow", "pillars_top10", "smells_top10",
        "magic_numbers_top20", "anomalies_top10", "domain_templates",
        "markers_to_fill",
    }
    assert expected_keys.issubset(set(dossier.keys()))

    # workbook_meta has the basics
    wm = dossier["workbook_meta"]
    assert wm["filename"] == SYNTH_PATH.name
    assert wm["sheet_count"] >= 1
    assert wm["vba_module_count"] >= 1
    assert isinstance(wm["complexity_total"], int)

    # markers_to_fill is non-empty and IDs match audit.md content
    md_text = (out / "audit.md").read_text(encoding="utf-8")
    marker_ids = {m["id"] for m in dossier["markers_to_fill"]}
    assert marker_ids, "no markers extracted"
    for mid in marker_ids:
        assert f"<!-- LLM-AUGMENT: {mid} -->" in md_text, (
            f"marker '{mid}' from dossier not present in audit.md"
        )

    # prompt.md contains every marker ID
    prompt_text = (out / "prompt.md").read_text(encoding="utf-8")
    for mid in marker_ids:
        assert mid in prompt_text, f"marker '{mid}' not in prompt.md"
    # And the prompt instructs strict-JSON output
    assert "JSON" in prompt_text
    assert "marker IDs" in prompt_text or "marker ID" in prompt_text


# ----------------------------------------------------------------------------
# 2. Ingest substitutes responses correctly
# ----------------------------------------------------------------------------

def test_ingest_substitutes_markers(tmp_path: Path):
    out = _run_harness_extract(tmp_path)
    dossier = json.loads((out / "dossier.json").read_text(encoding="utf-8"))

    # Build a fake responses.json with the first 3 markers
    sample_ids = [m["id"] for m in dossier["markers_to_fill"][:3]]
    fake_narratives = {
        mid: f"Test-narrative-for-{mid} (this string must appear in enriched output)."
        for mid in sample_ids
    }
    responses_path = out / "responses.json"
    responses_path.write_text(json.dumps(fake_narratives), encoding="utf-8")

    # Run ingest via CLI
    rc = cli_main([
        str(SYNTH_PATH), "--out-dir", str(out),
        "--ingest", str(responses_path),
    ])
    assert rc == 0
    enriched = (out / "audit-enriched.md").read_text(encoding="utf-8")

    # Each fake narrative appears in enriched output
    for mid, narrative in fake_narratives.items():
        assert narrative in enriched, (
            f"narrative for '{mid}' not substituted into audit-enriched.md"
        )
    # Markers themselves are still present (idempotent re-ingest)
    for mid in sample_ids:
        assert f"<!-- LLM-AUGMENT: {mid} -->" in enriched

    # HTML enriched output also exists and has the narratives
    enriched_html = (out / "audit-enriched.html").read_text(encoding="utf-8")
    for mid, narrative in fake_narratives.items():
        # HTML escapes content; narrative as written is plain ASCII so it
        # appears verbatim in the <p class="llm-narrative"> block.
        assert narrative in enriched_html, (
            f"narrative for '{mid}' not in audit-enriched.html"
        )


# ----------------------------------------------------------------------------
# 3. Graceful: missing markers in responses keeps heuristic narrative
# ----------------------------------------------------------------------------

def test_ingest_graceful_missing_marker(tmp_path: Path):
    out = _run_harness_extract(tmp_path)
    dossier = json.loads((out / "dossier.json").read_text(encoding="utf-8"))
    all_ids = [m["id"] for m in dossier["markers_to_fill"]]
    assert len(all_ids) >= 5, "synth should have several markers"

    # Only respond to the FIRST marker
    chosen = all_ids[0]
    fake = {chosen: "ONLY-NARRATIVE"}
    rp = out / "responses.json"
    rp.write_text(json.dumps(fake), encoding="utf-8")
    rc = cli_main([str(SYNTH_PATH), "--out-dir", str(out),
                    "--ingest", str(rp)])
    assert rc == 0
    enriched = (out / "audit-enriched.md").read_text(encoding="utf-8")
    original = (out / "audit.md").read_text(encoding="utf-8")

    # Chosen marker has the new narrative
    assert "ONLY-NARRATIVE" in enriched
    # Other markers still have heuristic content (we sample a few that
    # should always have something — vba-narration heuristics are non-empty).
    other_ids = [i for i in all_ids if i != chosen]
    # Pick one heuristic phrase and verify it survived for an un-replaced marker
    # If the heuristic phrase appeared only for `chosen` we couldn't tell;
    # but heuristics use phrases that recur across many sections.
    # Conservative check: enriched.md should still contain original sections
    # marked but unreplaced (each marker line still present).
    for mid in other_ids:
        assert f"<!-- LLM-AUGMENT: {mid} -->" in enriched
    # And heuristic-narrative phrasing present in original is also in enriched
    # (since most markers were unreplaced).
    # E.g. role-inference markers carry "**Role inference**:" lines.
    assert "**Role inference**" in enriched, (
        "heuristic VBA narration prose disappeared — graceful degradation broken"
    )
    # Sanity: original audit.md content should be MOSTLY preserved
    assert len(enriched) >= len(original) * 0.9


# ----------------------------------------------------------------------------
# 4. Graceful: extra response keys (not in audit.md) are ignored
# ----------------------------------------------------------------------------

def test_ingest_graceful_extra_keys(tmp_path: Path):
    out = _run_harness_extract(tmp_path)
    dossier = json.loads((out / "dossier.json").read_text(encoding="utf-8"))
    real_id = dossier["markers_to_fill"][0]["id"]
    fake = {
        real_id: "REAL-NARRATIVE",
        "vba-narration:DOES_NOT_EXIST": "garbage that should be ignored",
        "completely-unknown-marker:42": "another ignored",
    }
    rp = out / "responses.json"
    rp.write_text(json.dumps(fake), encoding="utf-8")
    rc = cli_main([str(SYNTH_PATH), "--out-dir", str(out),
                    "--ingest", str(rp)])
    assert rc == 0
    enriched = (out / "audit-enriched.md").read_text(encoding="utf-8")
    assert "REAL-NARRATIVE" in enriched
    assert "garbage that should be ignored" not in enriched
    assert "another ignored" not in enriched


# ----------------------------------------------------------------------------
# 5. Invalid JSON: ingest fails loudly with helpful message
# ----------------------------------------------------------------------------

def test_ingest_invalid_json(tmp_path: Path, capsys):
    out = _run_harness_extract(tmp_path)
    rp = out / "responses.json"
    rp.write_text("{this is not valid JSON,", encoding="utf-8")
    rc = cli_main([str(SYNTH_PATH), "--out-dir", str(out),
                    "--ingest", str(rp)])
    assert rc == 1
    err = capsys.readouterr().err
    # Error message must mention the file path AND the JSON error
    assert "FATAL" in err
    assert "JSON" in err.upper() or "json" in err
    assert "line" in err  # line number reported


def test_ingest_strips_markdown_fence(tmp_path: Path):
    """LLMs occasionally wrap JSON in ```json ... ``` — we tolerate that."""
    out = _run_harness_extract(tmp_path)
    dossier = json.loads((out / "dossier.json").read_text(encoding="utf-8"))
    real_id = dossier["markers_to_fill"][0]["id"]
    fake = {real_id: "FENCED-NARRATIVE"}
    rp = out / "responses.json"
    rp.write_text("```json\n" + json.dumps(fake) + "\n```\n", encoding="utf-8")
    rc = cli_main([str(SYNTH_PATH), "--out-dir", str(out),
                    "--ingest", str(rp)])
    assert rc == 0
    enriched = (out / "audit-enriched.md").read_text(encoding="utf-8")
    assert "FENCED-NARRATIVE" in enriched


# ----------------------------------------------------------------------------
# 6. Idempotency — extract twice produces byte-identical output
# ----------------------------------------------------------------------------

def test_harness_idempotency(tmp_path: Path):
    _require_synth()
    out1 = tmp_path / "run1"
    out2 = tmp_path / "run2"
    out1.mkdir(parents=True, exist_ok=True)
    out2.mkdir(parents=True, exist_ok=True)
    rc1 = cli_main([str(SYNTH_PATH), "--out-dir", str(out1), "--harness"])
    rc2 = cli_main([str(SYNTH_PATH), "--out-dir", str(out2), "--harness"])
    assert rc1 == 0 and rc2 == 0

    d1 = (out1 / "dossier.json").read_bytes()
    d2 = (out2 / "dossier.json").read_bytes()
    assert d1 == d2, "dossier.json differs between runs (idempotency broken)"

    p1 = (out1 / "prompt.md").read_bytes()
    p2 = (out2 / "prompt.md").read_bytes()
    assert p1 == p2, "prompt.md differs between runs (idempotency broken)"


# ----------------------------------------------------------------------------
# 7. Re-ingest is idempotent: ingesting an already-enriched md produces same
# ----------------------------------------------------------------------------

def test_reingest_idempotent(tmp_path: Path):
    """Running ingest twice on the same responses must produce byte-identical
    enriched output (the marker is preserved on first ingest so the second
    finds the same anchor)."""
    out = _run_harness_extract(tmp_path)
    dossier = json.loads((out / "dossier.json").read_text(encoding="utf-8"))
    sample_ids = [m["id"] for m in dossier["markers_to_fill"][:3]]
    fake = {mid: f"NARR-{mid}" for mid in sample_ids}
    rp = out / "responses.json"
    rp.write_text(json.dumps(fake), encoding="utf-8")

    rc1 = cli_main([str(SYNTH_PATH), "--out-dir", str(out),
                     "--ingest", str(rp)])
    assert rc1 == 0
    pass1 = (out / "audit-enriched.md").read_bytes()

    # Replace audit.md with the enriched copy and run ingest again.
    # Result should match pass1 byte-for-byte.
    (out / "audit.md").write_bytes(pass1)
    rc2 = cli_main([str(SYNTH_PATH), "--out-dir", str(out),
                     "--ingest", str(rp)])
    assert rc2 == 0
    pass2 = (out / "audit-enriched.md").read_bytes()
    assert pass1 == pass2, "re-ingest not idempotent"


# ----------------------------------------------------------------------------
# 8. Ingest fails gracefully when no audit.md exists yet
# ----------------------------------------------------------------------------

def test_ingest_requires_existing_audit(tmp_path: Path, capsys):
    _require_synth()
    out = tmp_path / "empty"
    out.mkdir(parents=True, exist_ok=True)
    rp = out / "responses.json"
    rp.write_text("{}", encoding="utf-8")
    rc = cli_main([str(SYNTH_PATH), "--out-dir", str(out),
                    "--ingest", str(rp)])
    assert rc == 1
    err = capsys.readouterr().err
    assert "audit.md" in err
    assert "FATAL" in err


# ----------------------------------------------------------------------------
# 9. Zero-LLM check: harness + ingest source files have no LLM imports
# ----------------------------------------------------------------------------

def test_harness_zero_llm_imports():
    """The whole architectural promise: harness/ingest never import an LLM
    SDK or make a network call. Pure scan over their source files."""
    import re
    src_dir = REPO_ROOT / "src" / "tier1_audit"
    forbidden = re.compile(
        r"\b(?:anthropic|openai|llm|gpt|cohere|langchain|llamaindex|"
        r"huggingface|transformers|requests|httpx|aiohttp|urllib3)\b",
        re.IGNORECASE,
    )
    offenders: list = []
    for fname in ("harness.py", "ingest.py"):
        f = src_dir / fname
        for lineno, line in enumerate(f.read_text(encoding="utf-8").splitlines(), 1):
            stripped = line.strip()
            if stripped.startswith(("#", '"', "'")):
                continue
            if stripped.startswith(("import ", "from ")) and forbidden.search(stripped):
                offenders.append(f"{fname}:{lineno}: {stripped}")
    assert not offenders, f"forbidden LLM/network imports in harness: {offenders}"


# ----------------------------------------------------------------------------
# 10. CLI --help mentions both new flags
# ----------------------------------------------------------------------------

def test_cli_help_mentions_harness_and_ingest(capsys):
    with pytest.raises(SystemExit):
        cli_main(["--help"])
    out = capsys.readouterr().out
    assert "--harness" in out
    assert "--ingest" in out


# ----------------------------------------------------------------------------
# 11. substitute_markers unit test (no CLI / disk)
# ----------------------------------------------------------------------------

def test_substitute_markers_unit():
    md = (
        "# Title\n\n"
        "## Section A\n\n"
        "<!-- LLM-AUGMENT: data-flow:Foo -->\n"
        "### `Foo` (visible, ...)\n"
        "**Role**: heuristic prose here.\n"
        "**Sources**: more heuristic prose.\n"
        "\n"
        "<!-- LLM-AUGMENT: vba-narration:Bar -->\n"
        "### `Bar` (data-loader, 100 lines)\n"
        "Bar heuristic narrative line 1.\n"
        "Bar heuristic narrative line 2.\n"
        "\n"
        "## Section B\n"
    )
    responses = {
        "data-flow:Foo": "Foo is the master sheet.",
        "vba-narration:Bar": "Bar implements the X algorithm.",
    }
    new_md, stats = substitute_markers(md, responses)
    assert "Foo is the master sheet." in new_md
    assert "Bar implements the X algorithm." in new_md
    # Heuristic prose must be removed for both
    assert "heuristic prose here" not in new_md
    assert "Bar heuristic narrative" not in new_md
    # Markers and H3 headings preserved
    assert "<!-- LLM-AUGMENT: data-flow:Foo -->" in new_md
    assert "### `Foo`" in new_md
    assert "<!-- LLM-AUGMENT: vba-narration:Bar -->" in new_md
    assert "### `Bar`" in new_md
    # H2 boundary preserved
    assert "## Section B" in new_md
    # Stats correct
    assert sorted(stats["replaced"]) == sorted(["data-flow:Foo", "vba-narration:Bar"])
    assert stats["kept_heuristic"] == []
    assert stats["unused_responses"] == []


def test_substitute_markers_dead_code_list_block():
    """The 'Possibly dead code' subsection has bullet items as heuristic content
    instead of H3-headed paragraphs. Substitution should replace the bullet."""
    md = (
        "### Possibly dead code (3 modules)\n\n"
        "<!-- LLM-AUGMENT: vba-narration:Foo.bas -->\n"
        "- `Foo.bas` (mixed, 62 lines, 1 subs)\n"
        "<!-- LLM-AUGMENT: vba-narration:Bar.cls -->\n"
        "- `Bar.cls` (data-loader, 31 lines, 2 subs)\n"
        "\n"
        "## Next Section\n"
    )
    new_md, stats = substitute_markers(md, {
        "vba-narration:Foo.bas": "Foo is unused legacy code.",
    })
    assert "Foo is unused legacy code." in new_md
    # Foo bullet line replaced
    assert "- `Foo.bas`" not in new_md
    # Bar bullet line preserved (no response for it)
    assert "- `Bar.cls`" in new_md
    assert "vba-narration:Bar.cls" in stats["kept_heuristic"]


# ----------------------------------------------------------------------------
# 12. load_responses validation
# ----------------------------------------------------------------------------

def test_load_responses_skips_invalid_entries(tmp_path: Path, capsys):
    rp = tmp_path / "r.json"
    rp.write_text(json.dumps({
        "valid:1": "good narrative",
        "valid:2": "",  # empty -> dropped
        "valid:3": 42,  # non-string -> dropped
    }), encoding="utf-8")
    out = load_responses(rp)
    assert out == {"valid:1": "good narrative"}
    err = capsys.readouterr().err
    assert "skipped" in err
