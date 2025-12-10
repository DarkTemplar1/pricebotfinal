#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
app_paths.py — Dynamiczne i przenośne ścieżki PriceBota.
"""

from __future__ import annotations
from pathlib import Path
import sys


def base_dir() -> Path:
    """
    Zwraca realny katalog, w którym działa program:

    - w EXE → folder obok PriceBot.exe
    - w dev → folder projektu (tam gdzie launcher_gui.py)
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys.argv[0]).resolve().parent
    else:
        return Path(__file__).resolve().parent.parent


def modules_dir() -> Path:
    """Folder z modułami (.py aktualizowalnymi)."""
    return base_dir() / "modules"


def data_file(name: str) -> Path:
    """Ścieżka do pliku danych obok EXE."""
    return base_dir() / name


TERYT_FILE = data_file("teryt.csv")
OBSZAR_SADOW_FILE = data_file("obszar_sadow.xlsx")
SLOWNIK_FILE = data_file("Słownik do Pricebota .xlsx")
POLSKA_FILE = data_file("Polska.xlsx")

TMP_DIR = base_dir() / "tmp"
LOGS_DIR = base_dir() / "logs"
EXPORTS_DIR = base_dir() / "exports"

for d in (TMP_DIR, LOGS_DIR, EXPORTS_DIR):
    d.mkdir(exist_ok=True)
