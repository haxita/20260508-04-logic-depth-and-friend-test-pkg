"""Allow `python -m tier1_audit ...` to invoke the CLI."""

from __future__ import annotations

import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())
