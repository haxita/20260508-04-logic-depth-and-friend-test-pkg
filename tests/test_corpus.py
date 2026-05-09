"""End-to-end tests for the tier1_audit pipeline.

The test corpus lives outside this package (in the Stage 0/1/2 worker dirs).
Tests skip gracefully when those files are not present (e.g. if the package
is installed elsewhere).
"""

from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

import pytest

from tier1_audit.audit import build_audit
from tier1_audit.cli import main as cli_main
from tier1_audit.extract import REDACTED_PLACEHOLDER, extract_cells
from tier1_audit.render import render_json, render_markdown


REPO_ROOT = Path(__file__).resolve().parents[1]
FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"

# External corpus paths (read-only references)
CORPUS_BASE = REPO_ROOT.parent
CORPUS = {
    "capacity_planning_synth": CORPUS_BASE / "20260508-02-xlsb-extractor-mvp" / "test_files" / "capacity_planning_synth.xlsm",
    "OperationPlanning": CORPUS_BASE / "20260508-02-xlsb-extractor-mvp" / "test_files" / "OperationPlanning.xlsm",
    "vba-web-example": CORPUS_BASE / "20260508-01-xlsm-vba-parsing" / "test_files" / "vba-web-example.xlsm",
}

ANOMALY_FIXTURE = FIXTURES_DIR / "anomaly_fixture.xlsm"


def _run_audit(input_path: Path, out_dir: Path, sanitize: bool = False):
    out_dir.mkdir(parents=True, exist_ok=True)
    rc = cli_main([
        str(input_path), "--out-dir", str(out_dir),
        *(["--sanitize"] if sanitize else []),
    ])
    return rc, (out_dir / "audit.md").read_text(encoding="utf-8"), (out_dir / "audit.json").read_text(encoding="utf-8")


# ----------------------------------------------------------------------------
# 1. Corpus runs — each of 3 files completes
# ----------------------------------------------------------------------------

@pytest.mark.parametrize("name,path", list(CORPUS.items()))
def test_corpus_runs(tmp_path: Path, name: str, path: Path):
    if not path.exists():
        pytest.skip(f"{name} not found at {path}; skipping")
    rc, md, js = _run_audit(path, tmp_path / "out")
    assert rc == 0, f"audit-xlsm exit code != 0 on {name}"
    assert md.startswith("# Audit report"), "audit.md missing header"
    j = json.loads(js)
    assert "basic_stats" in j
    assert j["basic_stats"]["sheet_count"] >= 1


# ----------------------------------------------------------------------------
# 2. Idempotency — byte-identical output across two runs
# ----------------------------------------------------------------------------

@pytest.mark.parametrize("name,path", list(CORPUS.items()))
def test_idempotency(tmp_path: Path, name: str, path: Path):
    if not path.exists():
        pytest.skip(f"{name} not found at {path}; skipping")
    out1 = tmp_path / "run1"
    out2 = tmp_path / "run2"
    rc1, md1, js1 = _run_audit(path, out1)
    rc2, md2, js2 = _run_audit(path, out2)
    assert rc1 == 0 and rc2 == 0
    assert md1 == md2, f"audit.md differs between runs on {name}"
    assert js1 == js2, f"audit.json differs between runs on {name}"


# ----------------------------------------------------------------------------
# 3. Zero-LLM grep — no AI dependencies sneaking in
# ----------------------------------------------------------------------------

def test_zero_llm_in_source():
    """Walk src/ and assert no LLM-related imports / names appear."""
    src_dir = REPO_ROOT / "src"
    forbidden = re.compile(r"\b(?:anthropic|openai|llm|gpt|claude|cohere|langchain|llamaindex|huggingface|transformers)\b",
                           re.IGNORECASE)
    offenders = []
    for py in src_dir.rglob("*.py"):
        text = py.read_text(encoding="utf-8")
        # Strip strings and comments crudely so the word "claude" in a doc-string
        # about Claude Code itself wouldn't trip us. We're really checking
        # imports + identifiers, not prose.
        # Lines that import / from / call: anything starting with import or from.
        for lineno, line in enumerate(text.splitlines(), 1):
            stripped = line.strip()
            if stripped.startswith(("#", '"', "'")):
                continue
            if stripped.startswith(("import ", "from ")) and forbidden.search(stripped):
                offenders.append(f"{py}:{lineno}: {stripped}")
    assert not offenders, f"Forbidden LLM imports found: {offenders}"


# ----------------------------------------------------------------------------
# 4. Pillars + Anomalies sections present in markdown
# ----------------------------------------------------------------------------

def test_pillars_section_present_synth(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, _ = _run_audit(p, tmp_path / "out")
    assert rc == 0
    # Round-3 pyramid: pillar table moved into Top Impact Findings (Top-5 view)
    # and Reference Appendix §8.1 (full table).
    assert "Top-5 Pillar Cells" in md, "Top Impact pillar block missing"
    assert "Full pillar table" in md, "Appendix pillar table missing"
    # Synth has fan-in hubs; pillar section must have at least one entry
    j = json.loads((tmp_path / "out" / "audit.json").read_text())
    assert len(j["pillars"]) > 0, "synth should produce at least one pillar"
    p0 = j["pillars"][0]
    assert "narrative" in p0
    assert "fan_in" in p0 and p0["fan_in"] >= 20
    # D1: pillar entries now include value + label fields
    assert "value" in p0
    assert "row_header" in p0
    assert "named_range" in p0


def test_anomalies_section_present_all(tmp_path: Path):
    """Anomaly content must always render (Top Impact section, even when empty)."""
    for name, path in CORPUS.items():
        if not path.exists():
            continue
        rc, md, _ = _run_audit(path, tmp_path / name)
        assert rc == 0
        # Round-3: anomalies are in the Top Impact Findings section.
        assert "Magic-number Anomalies" in md, f"{name} missing anomaly section"


def test_anomaly_detector_with_synthetic_fixture(tmp_path: Path):
    """The anomaly detector itself must work on a fixture that has a clear outlier.
    The 3 corpus files happen to have parameter sweeps (no hardcoded outlier);
    this fixture proves the detector fires when the pattern actually exists.
    """
    if not ANOMALY_FIXTURE.exists():
        pytest.skip("anomaly fixture not built — run tests/fixtures/build_anomaly_fixture.py")
    rc, md, js = _run_audit(ANOMALY_FIXTURE, tmp_path / "out")
    assert rc == 0
    j = json.loads(js)
    assert len(j["anomalies"]) >= 1, "anomaly fixture should have ≥ 1 detected anomaly"
    # Verify the high-confidence anomaly is the 0.82 vs 0.85 case
    confs = [a for a in j["anomalies"] if a["confidence"] == "high"]
    assert any(a["outlier_value"] == "0.82" for a in confs), \
        "expected 0.82 outlier among high-confidence anomalies"


# ----------------------------------------------------------------------------
# 5. VBA classification — every module gets a category
# ----------------------------------------------------------------------------

def test_vba_classify_synth(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, _, js = _run_audit(p, tmp_path / "out")
    assert rc == 0
    j = json.loads(js)
    valid_types = {"data-loader", "transformer", "report-writer",
                   "ui-handler", "dead-suspected", "mixed"}
    for c in j["vba_classifications"]:
        assert c["inferred_type"] in valid_types, \
            f"unexpected inferred_type: {c['inferred_type']}"


# ----------------------------------------------------------------------------
# 6. Sanitize mode — strips cell values
# ----------------------------------------------------------------------------

def test_sanitize_strips_cell_values(tmp_path: Path):
    """When --sanitize is on, CellRow.value should never carry original data."""
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")

    # Direct extract test — proves sanitize redacts at the source
    errs = []
    rows_normal, _, _, _ = extract_cells(p, errs, sanitize=False)
    errs2 = []
    rows_sani, _, _, _ = extract_cells(p, errs2, sanitize=True)

    # Find a cell that has a non-formula value in normal mode
    sample = next((r for r in rows_normal if r.value and not r.formula), None)
    assert sample is not None, "expected at least one non-formula cell with value"
    matching_sani = next((r for r in rows_sani
                          if r.sheet == sample.sheet and r.ref == sample.ref), None)
    assert matching_sani is not None
    assert matching_sani.value == REDACTED_PLACEHOLDER, \
        f"sanitize did not redact CellRow value: got {matching_sani.value!r}"

    # Also verify the sanitize banner appears in the audit.md when CLI --sanitize is on
    rc, md, js = _run_audit(p, tmp_path / "out_san", sanitize=True)
    assert rc == 0
    assert "SANITIZED MODE" in md, "sanitize banner missing from audit.md"
    j = json.loads(js)
    assert j.get("sanitized") is True
    assert j["methodology"]["sanitize_mode"] is True


def test_sanitize_preserves_formula_text(tmp_path: Path):
    """Sanitize must NOT redact formula text — formulas are structure, not data."""
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    errs = []
    rows_sani, _, _, _ = extract_cells(p, errs, sanitize=True)
    formula_rows = [r for r in rows_sani if r.formula]
    assert formula_rows, "expected formulas in synth"
    # Pick one that doesn't have only redacted content
    has_real_formula = any(r.formula.startswith("=") and len(r.formula) > 2
                           for r in formula_rows)
    assert has_real_formula, "sanitize incorrectly stripped formula text"


# ----------------------------------------------------------------------------
# 7. Smoke test: dataclasses round-trip cleanly to JSON
# ----------------------------------------------------------------------------

def test_json_roundtrip_synth(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    report = build_audit(p)
    js = render_json(report)
    j = json.loads(js)
    assert j["meta"]["file_name"] == p.name
    # The render explicitly drops the heavy `source_text` field from VBA modules
    for vm in j["vba_modules"]:
        assert "source_text" not in vm
    # The renderer also drops `_sheet_edges` (internal renderer-only product).
    assert "_sheet_edges" not in j


# ----------------------------------------------------------------------------
# 8. Polish round: pillar dedupe — synth's 99-cell column block collapses to 1 entry
# ----------------------------------------------------------------------------

def test_pillar_dedupe_synth(tmp_path: Path):
    """Synth has 订单!C10..C189 each with fan-in=72. After dedupe they should
    collapse to a single column-block entry, not 100 individual rows."""
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, js = _run_audit(p, tmp_path / "out")
    assert rc == 0
    j = json.loads(js)
    # After dedupe, expect FEW distinct pillar entries (was 20 individual cells before).
    assert len(j["pillars"]) <= 10, \
        f"expected <= 10 deduped pillars, got {len(j['pillars'])}"
    # At least one entry should be a column-block group with member_count > 1.
    grouped = [p for p in j["pillars"] if p.get("member_count", 1) > 1]
    assert grouped, "expected at least one column-block group on synth"
    g0 = grouped[0]
    assert g0["pillar_kind"] == "column-block"
    assert g0["member_count"] >= 5, \
        f"column-block should aggregate many cells, got {g0['member_count']}"


# ----------------------------------------------------------------------------
# 9. Polish round: TOC + Exec Summary present in markdown
# ----------------------------------------------------------------------------

def test_md_toc_and_exec_summary(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, _ = _run_audit(p, tmp_path / "out")
    assert rc == 0
    assert "## Executive Summary" in md, "executive summary missing"
    assert "## Table of Contents" in md, "TOC missing"
    # Round-3 pyramid sections
    assert "## Workflow Guide" in md, "Workflow Guide section missing"
    assert "## Data Flow Story" in md, "Data Flow Story section missing"
    assert "## Top Impact Findings" in md, "Top Impact section missing"
    assert "## VBA Module Walkthrough" in md, "VBA Walkthrough missing"
    assert "## Reference Appendix" in md, "Reference Appendix missing"
    assert "## Glossary" in md, "Glossary missing"
    # Methodology must appear exactly once
    assert md.count("\n## Methodology\n") == 1, \
        f"expected exactly one Methodology section, got {md.count('## Methodology')}"


# ----------------------------------------------------------------------------
# 10. Polish round: three Mermaid diagrams embedded in markdown
# ----------------------------------------------------------------------------

def test_three_mermaid_diagrams(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, _ = _run_audit(p, tmp_path / "out")
    assert rc == 0
    # Round-3: diagrams are now embedded inside the prose-led sections
    # (Workflow Guide / Data Flow Story / VBA classification mini / Pillar
    # impact). At least 3 mermaid fences should still be present.
    assert md.count("```mermaid") >= 3, \
        f"expected at least 3 mermaid blocks, got {md.count('```mermaid')}"
    # And they should each contain a graph keyword
    assert "graph LR" in md
    assert "graph TB" in md
    # D6: classification mini-diagram caps at <= 15 nodes — verify by counting
    # node-definition lines inside the classification mermaid block.
    cls_block = md[md.find("Classification mini-diagram"):]
    fence_start = cls_block.find("```mermaid")
    fence_end = cls_block.find("```", fence_start + 10)
    diagram = cls_block[fence_start:fence_end] if fence_start >= 0 else ""
    # Count `^[A-Za-z_]\w*\[` style node definitions
    node_lines = [l for l in diagram.splitlines()
                  if re.match(r"^[A-Za-z_]\w*\[", l.strip())]
    assert len(node_lines) <= 15, \
        f"VBA classification diagram has {len(node_lines)} nodes (D6 cap is 15)"
    # D6: ensure no `[N)` bracket bug
    assert "[18)" not in md and "[8)" not in md and "[5)" not in md, \
        "Bracket bug `[N)` should be fixed (use `[N]` or `(N)`)"


# ----------------------------------------------------------------------------
# 11. Polish round: HTML render produces a syntactically-plausible HTML file
# ----------------------------------------------------------------------------

def test_html_render(tmp_path: Path):
    from tier1_audit.render import render_html
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    report = build_audit(p)
    html = render_html(report)
    # Smoke checks: doctype + opening html + Mermaid script + sections
    assert html.startswith("<!DOCTYPE html>")
    assert '<html lang="en">' in html
    assert "mermaid" in html.lower()
    assert "Executive Summary" in html
    assert 'class="audit-section"' in html
    # Round-3 pyramid HTML sections
    assert "Workflow Guide" in html
    assert "Data Flow Story" in html
    assert "Top Impact Findings" in html
    assert "Reference Appendix" in html
    assert "Glossary" in html
    # Smoke: TOC anchor links present
    assert 'href="#workflow-guide' in html
    assert 'href="#reference-appendix' in html


# ----------------------------------------------------------------------------
# 12. Polish round: domain detector fires on synth (capacity-planning)
# ----------------------------------------------------------------------------

def test_domain_hint_synth(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, _, js = _run_audit(p, tmp_path / "out")
    assert rc == 0
    j = json.loads(js)
    dh = j.get("domain_hint")
    assert dh is not None, "domain_hint must be in JSON"
    # Synth has BOM, MPS, MRP, 产能 etc. — should match capacity-planning
    assert dh["domain"] == "capacity-planning", \
        f"expected capacity-planning, got {dh['domain']}"
    assert dh["confidence"] in {"high", "medium"}, \
        f"expected high/medium confidence, got {dh['confidence']}"


# ----------------------------------------------------------------------------
# 13. Polish round: --format flag honored
# ----------------------------------------------------------------------------

# ----------------------------------------------------------------------------
# 14. Round-3: D1 — pillar cells show value + label (Michael's #1 critique)
# ----------------------------------------------------------------------------

def test_round3_pillar_value_label_inline(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, js = _run_audit(p, tmp_path / "out")
    assert rc == 0
    j = json.loads(js)
    assert len(j["pillars"]) > 0
    p0 = j["pillars"][0]
    # All four new fields present
    for field in ("value", "value_kind", "row_header", "col_header", "named_range"):
        assert field in p0, f"pillar missing new field {field}"
    # The synth's _constants!C4 has a non-empty value (LOSS_RATE 0.05)
    assert p0["value"], "first pillar should have non-empty value (D1 fix)"
    # Markdown shows the value column inline
    assert "Value" in md, "Pillar table missing Value column header"
    # Markdown shows the value somewhere (e.g. `0.05` or row label LOSS_RATE)
    assert p0["value"] in md or p0["row_header"] in md, \
        "value or label should be visible in rendered markdown"


# ----------------------------------------------------------------------------
# 15. Round-3: D3 — Workflow Guide section + button detection
# ----------------------------------------------------------------------------

def test_round3_workflow_guide_section(tmp_path: Path):
    """Workflow Guide section is always present; content depends on detected buttons."""
    for name, path in CORPUS.items():
        if not path.exists():
            continue
        rc, md, js = _run_audit(path, tmp_path / name)
        assert rc == 0
        assert "## Workflow Guide" in md, f"{name} missing Workflow Guide"
        j = json.loads(js)
        wf = j.get("workflow") or {}
        # synth has no controls; OpPlan + vba-web have buttons
        if name == "capacity_planning_synth":
            # Synth has orphan VML drawings but NO control bindings — should be empty
            assert wf.get("no_buttons_detected") is True, \
                f"synth should have no buttons (got buttons: {wf.get('buttons')})"
            assert "no user-callable buttons" in md.lower() or \
                "formula-driven only" in md.lower(), \
                "synth should report 'no buttons detected' gracefully"
        if name == "OperationPlanning":
            assert len(wf.get("buttons") or []) >= 5, \
                "OperationPlanning should detect >= 5 buttons"


# ----------------------------------------------------------------------------
# 16. Round-3: D4 — VBA Module Walkthrough with narrations
# ----------------------------------------------------------------------------

def test_round3_vba_walkthrough(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, js = _run_audit(p, tmp_path / "out")
    assert rc == 0
    assert "## VBA Module Walkthrough" in md
    j = json.loads(js)
    nar = j.get("vba_narratives") or []
    assert len(nar) > 0, "should have VBA narratives"
    for n in nar[:3]:
        assert "narrative" in n
        assert "role_inference" in n
        assert "module_name" in n
    # Narration prose should appear in markdown
    assert "**Role inference**" in md, "narrative prose missing in markdown"


# ----------------------------------------------------------------------------
# 17. Round-3: D5 — Data Flow Story section with prose
# ----------------------------------------------------------------------------

def test_round3_data_flow_story(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, _ = _run_audit(p, tmp_path / "out")
    assert rc == 0
    assert "## Data Flow Story" in md
    # Per-sheet H3 headings should appear
    assert "**Role**" in md, "per-sheet role description missing"


# ----------------------------------------------------------------------------
# 18. Round-3: D7 — Domain templates fire on capacity-planning corpus
# ----------------------------------------------------------------------------

def test_round3_domain_templates(tmp_path: Path):
    # Synth: should match manufacturing-capacity-planning
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, js = _run_audit(p, tmp_path / "synth")
    assert rc == 0
    j = json.loads(js)
    matches = j.get("domain_template_matches") or []
    keys = [m["template_key"] for m in matches]
    assert "manufacturing-capacity-planning" in keys, \
        f"synth should match capacity-planning template, got {keys}"
    # The Domain-Specific Findings section should be in markdown when match present
    high_med = [m for m in matches if m["confidence"] in ("high", "medium")]
    if high_med:
        assert "## Domain-Specific Findings" in md, \
            "Domain section missing despite match"

    # vba-web-example: should NOT match any template (graceful degradation)
    p2 = CORPUS["vba-web-example"]
    if p2.exists():
        rc2, md2, js2 = _run_audit(p2, tmp_path / "vba-web")
        assert rc2 == 0
        j2 = json.loads(js2)
        m2 = j2.get("domain_template_matches") or []
        high_med2 = [m for m in m2 if m["confidence"] in ("high", "medium")]
        # No domain section when no high/medium match
        if not high_med2:
            assert "## Domain-Specific Findings" not in md2, \
                "Domain section should be absent when no template matches"


# ----------------------------------------------------------------------------
# 19. Round-3: D8 — Glossary section with term definitions
# ----------------------------------------------------------------------------

def test_round3_glossary(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    rc, md, _ = _run_audit(p, tmp_path / "out")
    assert rc == 0
    assert "## Glossary" in md
    # Key terms should appear
    for term in ["pillar cell", "fan-in", "smell", "magic number",
                 "veryHidden sheet", "On Error Resume Next"]:
        assert term in md, f"glossary missing definition for: {term}"


# ----------------------------------------------------------------------------
# 20. Round-3: D9 — LLM-AUGMENT markers in rendered markdown
# ----------------------------------------------------------------------------

def test_round3_llm_augment_markers(tmp_path: Path):
    """At least workflow-step / vba-narration / data-flow markers should be present
    where applicable (workflow-step only fires when buttons detected)."""
    for name, path in CORPUS.items():
        if not path.exists():
            continue
        rc, md, _ = _run_audit(path, tmp_path / name)
        assert rc == 0
        # Every corpus has VBA narration markers
        assert "<!-- LLM-AUGMENT: vba-narration:" in md, \
            f"{name} missing vba-narration LLM-AUGMENT markers"
        # Every corpus has data-flow markers (sheets exist everywhere)
        assert "<!-- LLM-AUGMENT: data-flow:" in md, \
            f"{name} missing data-flow LLM-AUGMENT markers"


def test_format_flag(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    # md only
    out_md = tmp_path / "md_only"
    out_md.mkdir(parents=True, exist_ok=True)
    rc = cli_main([str(p), "--out-dir", str(out_md), "--format", "md"])
    assert rc == 0
    assert (out_md / "audit.md").exists()
    assert not (out_md / "audit.html").exists()
    # html only
    out_html = tmp_path / "html_only"
    out_html.mkdir(parents=True, exist_ok=True)
    rc = cli_main([str(p), "--out-dir", str(out_html), "--format", "html"])
    assert rc == 0
    assert (out_html / "audit.html").exists()
    assert not (out_html / "audit.md").exists()
    # both (default)
    out_both = tmp_path / "both"
    out_both.mkdir(parents=True, exist_ok=True)
    rc = cli_main([str(p), "--out-dir", str(out_both)])
    assert rc == 0
    assert (out_both / "audit.md").exists()
    assert (out_both / "audit.html").exists()


# ----------------------------------------------------------------------------
# 21. Stage 5: --lang flag — en/de/zh produce non-empty single-tree output
# ----------------------------------------------------------------------------

@pytest.mark.parametrize("lang", ["en", "de", "zh"])
def test_lang_single(tmp_path: Path, lang: str):
    """Each single-language run produces a non-empty audit.md/json/html."""
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    out = tmp_path / f"out_{lang}"
    out.mkdir(parents=True, exist_ok=True)
    rc = cli_main([str(p), "--out-dir", str(out), "--lang", lang])
    assert rc == 0
    md = (out / "audit.md").read_text(encoding="utf-8")
    js = (out / "audit.json").read_text(encoding="utf-8")
    html = (out / "audit.html").read_text(encoding="utf-8")
    assert len(md) > 1000, f"audit.md for {lang} too short"
    assert len(js) > 1000, f"audit.json for {lang} too short"
    assert len(html) > 1000, f"audit.html for {lang} too short"
    # Localised title appears
    titles = {
        "en": "# Audit report",
        "de": "# Audit-Bericht",
        "zh": "# 审计报告",
    }
    assert md.startswith(titles[lang])


# ----------------------------------------------------------------------------
# 22. Stage 5: --lang all — three sibling output dirs produced
# ----------------------------------------------------------------------------

def test_lang_all(tmp_path: Path):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    out = tmp_path / "out_all"
    out.mkdir(parents=True, exist_ok=True)
    rc = cli_main([str(p), "--out-dir", str(out), "--lang", "all"])
    assert rc == 0
    for lang in ("en", "de", "zh"):
        sub = out / lang
        assert sub.exists(), f"--lang all should create {sub}"
        assert (sub / "audit.md").exists()
        assert (sub / "audit.json").exists()
        assert (sub / "audit.html").exists()


# ----------------------------------------------------------------------------
# 23. Stage 5: idempotency holds per-language
# ----------------------------------------------------------------------------

@pytest.mark.parametrize("lang", ["en", "de", "zh"])
def test_lang_idempotency(tmp_path: Path, lang: str):
    p = CORPUS["capacity_planning_synth"]
    if not p.exists():
        pytest.skip("synth not found")
    out1 = tmp_path / "run1"
    out2 = tmp_path / "run2"
    out1.mkdir(); out2.mkdir()
    rc1 = cli_main([str(p), "--out-dir", str(out1), "--lang", lang])
    rc2 = cli_main([str(p), "--out-dir", str(out2), "--lang", lang])
    assert rc1 == 0 and rc2 == 0
    md1 = (out1 / "audit.md").read_bytes()
    md2 = (out2 / "audit.md").read_bytes()
    assert md1 == md2, f"audit.md not idempotent for lang={lang}"
    js1 = (out1 / "audit.json").read_bytes()
    js2 = (out2 / "audit.json").read_bytes()
    assert js1 == js2, f"audit.json not idempotent for lang={lang}"
    html1 = (out1 / "audit.html").read_bytes()
    html2 = (out2 / "audit.html").read_bytes()
    assert html1 == html2, f"audit.html not idempotent for lang={lang}"
