"""CLI entry point — `audit-xlsm`.

Usage:
    audit-xlsm path/to/file.xlsm
    audit-xlsm path/to/file.xlsm --out-dir ./out --sanitize
    audit-xlsm path/to/file.xlsm --format html
    audit-xlsm path/to/file.xlsm --mermaid-inline
    audit-xlsm --help

Privacy:
    --sanitize  Replace every non-formula cell value with `<redacted>` in the
                report. Formulas, VBA source, smells, structure are preserved.
                Use this first to vet the report before sharing.
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
from .render import render_html, render_json, render_markdown


_MERMAID_CDN_URL = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"


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

    # JSON is always written.
    js = render_json(report)
    (out / "audit.json").write_text(js, encoding="utf-8")

    want_md = args.format in ("md", "both")
    want_html = args.format in ("html", "both")

    if want_md:
        md = render_markdown(report)
        (out / "audit.md").write_text(md, encoding="utf-8")

    if want_html:
        mermaid_inline_src = ""
        if args.mermaid_inline:
            try:
                mermaid_inline_src = _fetch_mermaid_js()
            except RuntimeError as e:
                print(f"WARN: --mermaid-inline failed; falling back to CDN: {e}",
                      file=sys.stderr)
                args.mermaid_inline = False
        html = render_html(
            report,
            mermaid_inline=args.mermaid_inline,
            mermaid_inline_source=mermaid_inline_src,
        )
        (out / "audit.html").write_text(html, encoding="utf-8")

    print(f"input    : {src.name}")
    print(f"out-dir  : {out}")
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
    return 0


if __name__ == "__main__":
    sys.exit(main())
