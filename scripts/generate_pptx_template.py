#!/usr/bin/env python3
"""Backward-compatible wrapper — delegates to the pptx_generator package.

This script exists so the Kiro Skill (SKILL.md) can reference
``scripts/generate_pptx_template.py`` without changes.

If the package is installed (pip install), it imports from the package.
Otherwise it falls back to running the generator module directly.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Add repo root to path so the package can be found even without pip install
_repo_root = Path(__file__).resolve().parent.parent
if str(_repo_root) not in sys.path:
    sys.path.insert(0, str(_repo_root))

from pptx_generator.generator import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
