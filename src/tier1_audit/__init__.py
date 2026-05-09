"""tier1_audit — Static-analysis audit pipeline for legacy xlsm workbooks.

Tier 1 of a tiered offering. Pure static analysis, zero LLM, zero network.
The headline capability is *logic comprehension*: pillar cells, magic-number
anomalies inside formula clusters, VBA module classification.

Public entry points:
    from tier1_audit import build_audit, render_markdown, render_json
    from tier1_audit.cli import main
"""

from __future__ import annotations

__version__ = "0.1.0"

from .audit import build_audit
from .render import render_markdown, render_json, render_html

__all__ = ["build_audit", "render_markdown", "render_json", "render_html", "__version__"]
