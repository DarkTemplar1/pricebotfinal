#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Launcher GUI – główne wejście do PriceBota.
Dynamicznie ładuje moduły z folderu "modules" i uruchamia selektor_csv.py.
Jest kompatybilny z PyInstaller.
"""

import sys
import os
import traceback
import importlib.util
from pathlib import Path
import tkinter as tk
from tkinter import messagebox


# ---------------------------------------------------------
# 1) Wymuszone importy dla PyInstaller (żeby znalazł zależności)
# ---------------------------------------------------------
try:
    import pandas as pd
    import numpy as np
    import openpyxl
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    pass


# ---------------------------------------------------------
# 2) Lokalizacja folderu "modules"
# ---------------------------------------------------------
def get_modules_dir() -> Path:
    """Zwraca ścieżkę do folderu modules obok PriceBot.exe"""
    if getattr(sys, 'frozen', False):
        # uruchomienie z .exe
        base = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(sys.executable).parent
        parent = Path(sys.executable).parent
    else:
        # uruchomienie z .py
        parent = Path(__file__).resolve().parent

    modules_path = parent / "modules"
    return modules_path


# ---------------------------------------------------------
# 3) Ładowanie modułu .py z pliku
# ---------------------------------------------------------
def load_module(path: Path):
    """Dynamicznie ładuje moduł Pythona z pliku."""
    if not path.exists():
        raise FileNotFoundError(f"Plik nie istnieje: {path}")

    spec = importlib.util.spec_from_file_location(path.stem, path)
    if not spec or not spec.loader:
        raise RuntimeError(f"Nie mogę załadować modułu: {path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# ---------------------------------------------------------
# 4) Główna logika startowa
# ---------------------------------------------------------
def main():
    modules_dir = get_modules_dir()

    # upewniamy się, że folder istnieje
    if not modules_dir.exists():
        messagebox.showerror(
            "PriceBot – błąd startu",
            f"Nie znaleziono folderu MODULES:\n{modules_dir}\n\n"
            "Upewnij się, że katalog 'modules' leży obok PriceBot.exe."
        )
        return

    # szukamy selektor_csv.py
    selector_path = modules_dir / "selektor_csv.py"

    if not selector_path.exists():
        messagebox.showerror(
            "PriceBot – błąd startu",
            f"Nie znaleziono pliku GUI:\n{selector_path}\n\n"
            "Upewnij się, że katalog 'modules' zawiera selektor_csv.py."
        )
        return

    try:
        selector = load_module(selector_path)
    except Exception as e:
        tb = traceback.format_exc()
        messagebox.showerror(
            "Błąd uruchamiania PriceBota",
            f"{e}\n\nŚlad błędu:\n{tb}"
        )
        return

    # uruchamiamy klasę App z selektor_csv.py
    if hasattr(selector, "main"):
        try:
            selector.main()
        except Exception as e:
            tb = traceback.format_exc()
            messagebox.showerror(
                "Błąd w trakcie działania",
                f"{e}\n\nŚlad:\n{tb}"
            )
    else:
        messagebox.showerror(
            "Błąd",
            "Plik selektor_csv.py nie posiada funkcji main()."
        )


# ---------------------------------------------------------
# 5) Start
# ---------------------------------------------------------
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        tb = traceback.format_exc()
        messagebox.showerror("Launcher – krytyczny błąd", f"{e}\n\n{tb}")
