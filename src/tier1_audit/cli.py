"""CLI entry point — `audit-xlsm`.

Usage:
    audit-xlsm path/to/file.xlsm
    audit-xlsm path/to/file.xlsm --out-dir ./out --sanitize
    audit-xlsm path/to/file.xlsm --format html
    audit-xlsm path/to/file.xlsm --mermaid-inline
    audit-xlsm path/to/file.xlsm --harness         # Track B: emit dossier+prompt
    audit-xlsm path/to/file.xlsm --ingest responses.json
    audit-xlsm --help

Privacy:
    --sanitize  Replace every non-formula cell value with `<redacted>` in the
                report. Formulas, VBA source, smells, structure are preserved.
                Use this first to vet the report before sharing.

Track A vs Track B (see docs/harness-guide.md):
    Track A (default): zero-LLM static analysis. Outputs audit.{md,html,json}.
    Track B (--harness + --ingest): user pipes the prompt through their OWN
    LLM client (Copilot Chat, Claude Desktop, ChatGPT) and pastes the JSON
    back. The tool itself never calls an LLM.
"""

from __future__ import annotations

import argparse
import sys
import traceback
import urllib.error
import urllib.request
from pathlib import Path

from . import __version__
from .audit import build_audit
from .harness import extract as harness_extract
from .i18n import SUPPORTED_LANGS
from .ingest import ingest as harness_ingest
from .render import render_html, render_json, render_markdown


_MERMAID_CDN_URL = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"
_LANG_CHOICES = ("en", "de", "zh", "all")


def _build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="audit-xlsm",
        description=(
            "Tier 1 audit pipeline for legacy xlsm workbooks. "
            "Produces a markdown + html + json report with: pillar cells, "
            "magic-number anomalies, VBA module classification, Hermans-style "
            "smells, and complexity score. Pure static analysis, zero LLM, "
            "zero network at audit time."
        ),
    )
    ap.add_argument("input", nargs="?",
                    help="path to .xlsm file (or use --input)")
    ap.add_argument("--input", dest="input_flag",
                    help="alternate flag form for the input file path")
    ap.add_argument("--out-dir", default="./out",
                    help="output directory (default: ./out, created if missing)")
    ap.add_argument("--sanitize", action="store_true",
                    help="redact every non-formula cell value in the output "
                         "(formulas + structure preserved; safe for sharing)")
    ap.add_argument("--format", choices=["md", "html", "both"], default="both",
                    help="output formats to write (default: both)")
    ap.add_argument("--mermaid-inline", action="store_true",
                    help="download mermaid.min.js once at audit time and inline "
                         "it into the HTML for fully offline viewing "
                         "(default: reference public CDN at view time)")
    ap.add_argument("--harness", action="store_true",
                    help="Track B: in addition to audit.{md,html,json}, also "
                         "emit dossier.json + prompt.md so the user can paste "
                         "them into their own LLM client (Copilot/Claude/etc.) "
                         "and feed the response back via --ingest. "
                         "We never call an LLM ourselves.")
    ap.add_argument("--ingest", metavar="RESPONSES_JSON", default=None,
                    help="Track B ingest: read the LLM's JSON response and "
                         "produce audit-enriched.{md,html} alongside the "
                         "originals. Path is to a JSON file the user saved "
                         "after pasting the harness prompt into their LLM.")
    ap.add_argument("--lang", choices=list(_LANG_CHOICES), default="en",
                    help="Output language: en (default, master), de (German), "
                         "zh (Chinese), or all (produce three sibling output "
                         "directories <out>/en, <out>/de, <out>/zh). "
                         "English is canonical; DE/ZH translations use "
                         "SAP/manufacturing-industry-standard terminology.")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return ap


def _fetch_mermaid_js() -> str:
    """Fetch mermaid.min.js from the CDN. Only invoked when --mermaid-inline is set."""
    try:
        with urllib.request.urlopen(_MERMAID_CDN_URL, timeout=30) as resp:
            return resp.read().decode("utf-8")
    except (urllib.error.URLError, TimeoutError) as e:
        raise RuntimeError(
            f"failed to fetch mermaid.min.js for --mermaid-inline: {type(e).__name__}: {e}"
        ) from e


def main(argv=None) -> int:
    ap = _build_parser()
    args = ap.parse_args(argv)

    input_path_str = args.input or args.input_flag
    if not input_path_str:
        ap.error("input file is required (positional argument or --input)")

    src = Path(input_path_str).expanduser().resolve()
    out = Path(args.out_dir).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)

    if not src.exists():
        print(f"FATAL: input file not found: {src}", file=sys.stderr)
        return 1
    if not src.is_file():
        print(f"FATAL: input path is not a file: {src}", file=sys.stderr)
        return 1

    # Pure-ingest path: --ingest is set and audit.md already exists in out-dir.
    # Skip the (slow) audit rebuild — just substitute markers.
    if args.ingest:
        try:
            stats_summary = harness_ingest(
                audit_dir=out,
                responses_path=Path(args.ingest).expanduser().resolve(),
                lang=args.lang,
            )
        except SystemExit as e:
            print(str(e), file=sys.stderr)
            return 1
        except FileNotFoundError as e:
            print(f"FATAL: {e}", file=sys.stderr)
            return 1
        print(f"input        : {src.name}")
        print(f"out-dir      : {out}")
        print(f"responses    : {args.ingest}")
        print(f"lang         : {args.lang}")
        for entry in stats_summary["entries"]:
            print(f"enriched md  : {entry['md_path']} [{entry['lang']}]")
            if entry.get("html_path"):
                print(f"enriched html: {entry['html_path']} [{entry['lang']}]")
            print(f"  replaced   : {len(entry['md_replaced'])} marker(s)")
            print(f"  kept (h.)  : {len(entry['md_kept_heuristic'])} marker(s) "
                  f"with no LLM response — heuristic narrative kept")
            if entry["md_unused_responses"]:
                print(f"  unused     : {len(entry['md_unused_responses'])} response key(s) "
                      f"didn't match any marker (ignored)")
        return 0

    try:
        report = build_audit(src, sanitize=args.sanitize)
    except Exception as e:
        tb = traceback.format_exc()
        msg = (
            f"# Audit failed — `{src.name}`\n\n"
            f"```\n{type(e).__name__}: {e}\n\n{tb}\n```\n"
        )
        (out / "audit-failed.md").write_text(msg, encoding="utf-8")
        print(f"FATAL: audit pipeline crashed: {type(e).__name__}: {e}", file=sys.stderr)
        return 1

    want_md = args.format in ("md", "both")
    want_html = args.format in ("html", "both")

    # Resolve target languages
    if args.lang == "all":
        target_langs = list(SUPPORTED_LANGS)
    else:
        target_langs = [args.lang]

    # When --lang all, write to <out>/<lang>/audit.* — three sibling trees.
    # When single lang, write to <out>/audit.* (current behavior).
    def _resolve_lang_outdir(lang: str) -> Path:
        if args.lang == "all":
            d = out / lang
            d.mkdir(parents=True, exist_ok=True)
            return d
        return out

    mermaid_inline_src = ""
    if want_html and args.mermaid_inline:
        try:
            mermaid_inline_src = _fetch_mermaid_js()
        except RuntimeError as e:
            print(f"WARN: --mermaid-inline failed; falling back to CDN: {e}",
                  file=sys.stderr)
            args.mermaid_inline = False

    harness_marker_count = None
    for lang in target_langs:
        lang_dir = _resolve_lang_outdir(lang)
        # JSON is always written.
        js = render_json(report, lang=lang)
        (lang_dir / "audit.json").write_text(js, encoding="utf-8")
        if want_md:
            md = render_markdown(report, lang=lang)
            (lang_dir / "audit.md").write_text(md, encoding="utf-8")
        if want_html:
            html = render_html(
                report,
                mermaid_inline=args.mermaid_inline,
                mermaid_inline_source=mermaid_inline_src,
                lang=lang,
            )
            (lang_dir / "audit.html").write_text(html, encoding="utf-8")
        # Track B harness extract: emit dossier.json + prompt.md alongside
        # the standard audit.* outputs. When --lang all, the harness writes
        # ONE dossier+prompt pair per language directory (so the LLM-generated
        # narratives match the language's audit.md marker layout).
        if args.harness:
            if want_md:
                md_for_harness = (lang_dir / "audit.md").read_text(encoding="utf-8")
            else:
                md_for_harness = render_markdown(report, lang=lang)
            try:
                h_stats = harness_extract(
                    report=report,
                    audit_md_text=md_for_harness,
                    out_dir=lang_dir,
                    source_path=src,
                    lang=args.lang,  # tells the prompt builder whether to
                                     # request trilingual JSON or single-lang
                )
                harness_marker_count = h_stats["marker_count"]
            except Exception as e:
                tb = traceback.format_exc()
                print(f"WARN: --harness extract failed: {type(e).__name__}: {e}\n{tb}",
                      file=sys.stderr)
                harness_marker_count = None

    print(f"input    : {src.name}")
    print(f"out-dir  : {out}")
    print(f"lang     : {args.lang}")
    print(f"format   : {args.format}{' + mermaid-inline' if args.mermaid_inline and want_html else ''}")
    print(f"sanitize : {'on' if args.sanitize else 'off'}")
    print(f"complex. : {report.complexity.total}/100")
    print(f"smells   : {len(report.smells)}")
    print(f"pillars  : {len(report.pillars)}")
    print(f"anomalies: {len(report.anomalies)}")
    print(f"vba clf  : {len(report.vba_classifications)} module(s) classified")
    if report.domain_hint is not None:
        dh = report.domain_hint
        print(f"domain   : {dh.domain} (confidence: {dh.confidence}, hits: {dh.hits})")
    print(f"sheets   : {report.basic_stats.sheet_count}")
    print(f"formulas : {report.basic_stats.cell_count_formula}")
    print(f"vba      : {report.basic_stats.vba_module_count} modules / "
          f"{report.basic_stats.vba_total_lines} lines")
    print(f"errors   : {report.basic_stats.parse_errors_count}")
    if args.harness and harness_marker_count is not None:
        print(f"harness  : dossier.json + prompt.md "
              f"({harness_marker_count} marker(s) to fill, lang={args.lang})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
