#!/usr/bin/env python3
"""mTails - a modern, Excel-compatible spreadsheet.

Entry point kept at the repository root so the application can be started with
``python3 main.py`` on any platform that ships PySide6.
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from mtails.app import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
