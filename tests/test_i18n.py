"""Tests for the i18n catalog + loader (Stage 5).

Architectural promises being verified:
  - Three locale catalogs exist (en, de, zh) with EXACTLY the same key set.
  - The `t()` function looks up keys, formats with vars, and falls back to
    English when a key is missing in a target language.
  - When a key is missing in BOTH target and English, t() returns
    `[[missing:KEY]]` so the caller sees the bug.
  - Trilingual rendering through the public render functions returns
    non-empty output for each language.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from tier1_audit.i18n import (
    DEFAULT_LANG,
    SUPPORTED_LANGS,
    load_catalog,
    md_table_separator,
    reset_cache,
    split_pipe_columns,
    t,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
I18N_DIR = REPO_ROOT / "src" / "tier1_audit" / "i18n"


@pytest.fixture(autouse=True)
def _reset():
    """Each test starts with a clean catalog cache."""
    reset_cache()
    yield
    reset_cache()


# ----------------------------------------------------------------------------
# 1. Catalog files exist for en/de/zh
# ----------------------------------------------------------------------------

def test_catalog_files_exist():
    for lang in SUPPORTED_LANGS:
        p = I18N_DIR / f"{lang}.json"
        assert p.exists(), f"i18n catalog {lang}.json missing"
        # Parses as JSON and is a dict
        data = json.loads(p.read_text(encoding="utf-8"))
        assert isinstance(data, dict)
        assert len(data) >= 80, f"{lang}.json should have >= 80 keys"


# ----------------------------------------------------------------------------
# 2. Every key in en.json has a matching key in de.json AND zh.json
# ----------------------------------------------------------------------------

def test_catalog_key_parity():
    en = json.loads((I18N_DIR / "en.json").read_text(encoding="utf-8"))
    de = json.loads((I18N_DIR / "de.json").read_text(encoding="utf-8"))
    zh = json.loads((I18N_DIR / "zh.json").read_text(encoding="utf-8"))
    en_keys = set(en.keys())
    de_keys = set(de.keys())
    zh_keys = set(zh.keys())
    missing_in_de = sorted(en_keys - de_keys)
    extra_in_de = sorted(de_keys - en_keys)
    missing_in_zh = sorted(en_keys - zh_keys)
    extra_in_zh = sorted(zh_keys - en_keys)
    assert not missing_in_de, f"keys missing in de.json: {missing_in_de}"
    assert not extra_in_de, f"unknown keys in de.json: {extra_in_de}"
    assert not missing_in_zh, f"keys missing in zh.json: {missing_in_zh}"
    assert not extra_in_zh, f"unknown keys in zh.json: {extra_in_zh}"
    # And the catalog size is in the right range
    assert 80 <= len(en_keys) <= 400, (
        f"en.json catalog has {len(en_keys)} keys; expected 80-400"
    )


# ----------------------------------------------------------------------------
# 3. t() lookups + var formatting work for every language
# ----------------------------------------------------------------------------

def test_t_basic_lookup_each_lang():
    # Every language should have exec_summary.heading
    for lang in SUPPORTED_LANGS:
        s = t("exec_summary.heading", lang)
        assert s, f"empty value for exec_summary.heading in {lang}"
        assert "[[missing" not in s


def test_t_format_with_vars():
    s = t("exec_summary.complexity_line", "en", total=42, tier="moderate")
    assert "42" in s
    assert "moderate" in s
    s_de = t("exec_summary.complexity_line", "de", total=42, tier="mittel")
    assert "42" in s_de
    assert "mittel" in s_de


# ----------------------------------------------------------------------------
# 4. Fallback to English when a key is missing in target language
# ----------------------------------------------------------------------------

def test_fallback_to_english_when_missing(tmp_path, monkeypatch):
    """If we add a key only to en.json (but not to de.json), t('key', 'de')
    should fall back to the EN text. We simulate this by patching the in-
    memory catalog cache."""
    # Load the real catalogs first
    load_catalog("en")
    load_catalog("de")
    # Inject a fake key into en cache
    from tier1_audit import i18n as i18n_module
    i18n_module._CATALOG_CACHE["en"]["fake.test_only"] = "Hello {name}"
    # de catalog doesn't have it — should fall back
    s = t("fake.test_only", "de", name="World")
    assert s == "Hello World"


def test_missing_in_both_returns_marker():
    s = t("definitely.does.not.exist.anywhere", "en")
    assert s == "[[missing:definitely.does.not.exist.anywhere]]"
    s = t("definitely.does.not.exist.anywhere", "de")
    assert s == "[[missing:definitely.does.not.exist.anywhere]]"


# ----------------------------------------------------------------------------
# 5. Glossary terms — ensure key terminology was translated (not English
#    leftovers in DE/ZH — the most common bug)
# ----------------------------------------------------------------------------

def test_de_glossary_uses_german_terminology():
    en = t("glossary.fan_in", "en")
    de = t("glossary.fan_in", "de")
    # English fan-in entry contains 'distinct formulas reference'
    assert "distinct formulas" in en
    # German entry should contain Eingangsgrad
    assert "Eingangsgrad" in de
    # And the glossary entry for "Stückliste" / Bill of Materials concepts
    # appears via the column-block term — not specific assertion needed there.


def test_zh_glossary_uses_chinese_terminology():
    en = t("glossary.pillar", "en")
    zh = t("glossary.pillar", "zh")
    assert "high fan-in" in en
    # Chinese version uses "支柱单元格" or "引用数"
    assert ("支柱" in zh) or ("引用数" in zh)


# ----------------------------------------------------------------------------
# 6. Helpers: split_pipe_columns + md_table_separator
# ----------------------------------------------------------------------------

def test_split_pipe_columns_and_separator():
    cols = split_pipe_columns(t("pillar.table_columns", "en"))
    assert len(cols) == 9
    assert cols[0] == "Rank"
    sep = md_table_separator(len(cols))
    assert sep == "|---|---|---|---|---|---|---|---|---|"


# ----------------------------------------------------------------------------
# 7. The English audit output is byte-identical pre/post i18n refactor
#    (sanity: smoke-test that render_markdown('en') still produces something
#    starting with the right cover heading).
# ----------------------------------------------------------------------------

def test_render_markdown_en_starts_correctly():
    from tier1_audit.audit import build_audit
    from tier1_audit.render import render_markdown
    p = REPO_ROOT.parent / "20260508-02-xlsb-extractor-mvp" / "test_files" / "capacity_planning_synth.xlsm"
    if not p.exists():
        pytest.skip(f"synth corpus not present at {p}")
    report = build_audit(p)
    md = render_markdown(report, lang="en")
    assert md.startswith("# Audit report — `capacity_planning_synth.xlsm`")
    # And the German render starts with the German title
    md_de = render_markdown(report, lang="de")
    assert md_de.startswith("# Audit-Bericht — `capacity_planning_synth.xlsm`")
    # Chinese
    md_zh = render_markdown(report, lang="zh")
    assert md_zh.startswith("# 审计报告 — `capacity_planning_synth.xlsm`")
