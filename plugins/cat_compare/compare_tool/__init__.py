"""
__init__.py
-----------
This file initializes the `compare_tool` package.

Purpose:
- Marks the directory as a Python package.
- Exposes key modules and functions for external use.
"""

# compare_tool/__init__.py

from .config import load_config
from .service import run_comparison
