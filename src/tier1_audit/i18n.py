"""i18n catalog loader for trilingual rendering (en / de / zh).

Architectural promise: zero new deps. Pure stdlib (json + str.format).

Catalog files live at `src/tier1_audit/i18n/<lang>.json`. Each is a flat
dict of dot-namespaced string keys to f-string-like templates with `{var}`
placeholders. English (`en.json`) is the master catalog; `de.json` and
`zh.json` are derived translations.

Usage:
    from .i18n import t, load_catalog, SUPPORTED_LANGS
    s = t("pillar.heading", "de")
    s = t("workflow.step_line", "zh", order=1, label="Run", sheet="Main",
          module_name="Mod", sub_name="Calc")

Fallback chain (load_catalog applies it as the catalog is constructed):
    1. Look in target language catalog
    2. If missing, fall back to English (logged once per missing key)
    3. If missing in English too, return literal `[[missing:KEY]]`
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Optional

SUPPORTED_LANGS = ("en", "de", "zh")
DEFAULT_LANG = "en"

_CATALOG_DIR = Path(__file__).parent / "i18n"

# Cache for loaded catalogs: {lang: {key: template_string}}
_CATALOG_CACHE: dict = {}

# Track which (lang, key) misses have already been warned about, so we don't
# spam stderr per render call.
_WARNED_MISSING: set = set()


def _catalog_path(lang: str) -> Path:
    return _CATALOG_DIR / f"{lang}.json"


def _read_raw_catalog(lang: str) -> dict:
    """Read the raw JSON catalog for a language. Raises FileNotFoundError if
    the file isn't present.
    """
    p = _catalog_path(lang)
    if not p.exists():
        raise FileNotFoundError(f"i18n catalog not found: {p}")
    with open(p, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError(f"i18n catalog must be a JSON object: {p}")
    return data


def load_catalog(lang: str) -> dict:
    """Return the catalog dict for `lang`, loading it once and caching.

    Always loads English as well so the fallback chain works deterministically
    without re-reading from disk. The returned catalog is the language's own
    map; callers should use `t()` (which applies fallback) rather than reading
    keys from the dict directly.
    """
    if lang not in SUPPORTED_LANGS:
        raise ValueError(
            f"unsupported lang: {lang!r} (supported: {SUPPORTED_LANGS})"
        )
    if lang in _CATALOG_CACHE:
        return _CATALOG_CACHE[lang]

    # Always pre-load English (fallback target) too.
    if DEFAULT_LANG not in _CATALOG_CACHE:
        _CATALOG_CACHE[DEFAULT_LANG] = _read_raw_catalog(DEFAULT_LANG)

    if lang == DEFAULT_LANG:
        return _CATALOG_CACHE[DEFAULT_LANG]

    _CATALOG_CACHE[lang] = _read_raw_catalog(lang)
    return _CATALOG_CACHE[lang]


def _warn_once(lang: str, key: str, kind: str) -> None:
    sig = (lang, key, kind)
    if sig in _WARNED_MISSING:
        return
    _WARNED_MISSING.add(sig)
    print(
        f"WARN: i18n: key {key!r} {kind} for lang={lang!r} — "
        f"falling back to {DEFAULT_LANG!r}",
        file=sys.stderr,
    )


def t(key: str, lang: str = DEFAULT_LANG, **vars) -> str:
    """Look up `key` in the lang catalog, format with `**vars`, fall back
    to English when missing.

    If `key` is missing in both target lang AND English, returns
    `[[missing:KEY]]` so it's obvious in output.

    Uses `str.format(**vars)` for substitution — KISS, no Jinja. Unused vars
    are silently ignored. Missing vars raise KeyError (caller should pass
    them all; this is intentional — silent failure here would mask bugs).
    """
    if lang not in SUPPORTED_LANGS:
        # Defensive: callers should pass a validated lang, but if they
        # don't, fall through to English instead of raising.
        lang = DEFAULT_LANG

    cat = load_catalog(lang)
    en_cat = load_catalog(DEFAULT_LANG)

    template = cat.get(key)
    if template is None:
        # Try English fallback
        if lang != DEFAULT_LANG:
            _warn_once(lang, key, "missing")
        template = en_cat.get(key)

    if template is None:
        # Both target and English lack this key — emit a visible marker
        return f"[[missing:{key}]]"

    if not vars:
        return template
    try:
        return template.format(**vars)
    except KeyError as e:
        # Missing format var — emit a visible marker so the caller sees
        # the bug rather than a silent partial substitution.
        return f"[[i18n-format-error:{key}:missing-var:{e.args[0]}]]"


def reset_cache() -> None:
    """Clear the catalog cache. Test-only utility."""
    _CATALOG_CACHE.clear()
    _WARNED_MISSING.clear()


def list_catalog_keys(lang: str = DEFAULT_LANG) -> list:
    """Return the sorted list of keys in `lang`'s catalog (loading if needed)."""
    return sorted(load_catalog(lang).keys())


def split_pipe_columns(value: str) -> list:
    """Split a `Col1|Col2|Col3` column-name string back into a list.

    Used for table-header keys whose value uses `|` as a column separator
    (so DE / ZH translators only translate the column names, not the
    Markdown table boilerplate).
    """
    return [c.strip() for c in value.split("|")]


def md_table_separator(n_cols: int) -> str:
    """Build a markdown table separator row of `n_cols` columns.

    Returns `"|---|---|...|"`. Used together with split_pipe_columns to
    rebuild full Markdown table headers from a translated column-name string.
    """
    return "|" + "---|" * n_cols
