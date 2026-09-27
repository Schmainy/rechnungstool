"""Legt den Projektordner auf sys.path, damit `from src import ...` klappt."""

import sys
from pathlib import Path

PROJEKT_DIR = Path(__file__).resolve().parent.parent
if str(PROJEKT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJEKT_DIR))
