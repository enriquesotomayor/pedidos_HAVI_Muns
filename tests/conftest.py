# -*- coding: utf-8 -*-
"""Hace importable havi2odoo (raíz del repo) desde los tests."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
