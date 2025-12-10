#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
CzyszczenieAdresu.py – uzupełnianie dziur adresowych:
Województwo, Powiat, Gmina, Miejscowość, Dzielnica.

Źródła danych:
- teryt.csv          (opcjonalny)
- obszar_sadow.xlsx  (opcjonalny)

Braki = puste / NaN / '---'
"""

from __future__ import annotations

import argparse
from pathlib import Path
import unicodedata
import pandas as pd
import math
import traceback

ADDR_COLS = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]


# ------------------- HELPERY -------------------

def _norm(s: str) -> str:
    """Normalizacja do dopasowań."""
    s = str(s or "").strip().lower()
    s = "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )
    while "  " in s:
        s = s.replace("  ", " ")
    return s


def _is_missing(v) -> bool:
    """Czy traktujemy to jako brak? (puste / NaN / '---')."""
    if v is None:
        return True
    if isinstance(v, float):
        try:
            if math.isnan(v):
                return True
        except Exception:
            pass
    s = str(v).strip()
    return s == "" or s == "---"


# ------------------- Wczytanie źródeł -------------------

def load_teryt(path: str = "teryt.csv") -> pd.DataFrame:
    """
    Wczytuje teryt.csv i przygotowuje kolumny znormalizowane.
    Jeśli plik nie istnieje – zwraca pustą ramkę i NIE traktuje tego jako błąd.
    """
    p = Path(path)
    if not p.exists():
        print(f"[INFO] teryt.csv nie znaleziony – pomijam TERYT.")
        return pd.DataFrame(columns=[
            "Wojewodztwo", "Powiat", "Gmina", "Miejscowosc", "Dzielnica",
            "woj_n", "pow_n", "gmi_n", "miej_n", "dz_n"
        ])

    df = pd.read_csv(p, sep=";", dtype=str).fillna("")

    required = ["Wojewodztwo", "Powiat", "Gmina", "Miejscowosc", "Dzielnica"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Brak kolumny '{col}' w teryt.csv")

    for col in required:
        df[col] = df[col].astype(str)

    df["woj_n"]  = df["Wojewodztwo"].map(_norm)
    df["pow_n"]  = df["Powiat"].map(_norm)
    df["gmi_n"]  = df["Gmina"].map(_norm)
    df["miej_n"] = df["Miejscowosc"].map(_norm)
    df["dz_n"]   = df["Dzielnica"].map(_norm)

    return df


def load_obszar_sadow(path: str = "obszar_sadow.xlsx") -> pd.DataFrame:
    """
    Wczytuje obszar_sadow.xlsx i przygotowuje kolumny znormalizowane.
    Jeśli plik nie istnieje – zwraca pustą ramkę i NIE traktuje tego jako błąd.
    """
    p = Path(path)
    if not p.exists():
        print(f"[INFO] obszar_sadow.xlsx nie znaleziony – pomijam źródło sądów.")
        return pd.DataFrame(columns=[
            "Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica",
            "woj_n", "pow_n", "gmi_n", "miej_n", "dz_n"
        ])

    df = pd.read_excel(p, dtype=str).fillna("")

    required = ["Województwo", "Powiat", "Gmina", "Miejscowość", "Dzielnica"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Brak kolumny '{col}' w pliku {p}")

    for col in required:
        df[col] = df[col].astype(str)

    df["woj_n"]  = df["Województwo"].map(_norm)
    df["pow_n"]  = df["Powiat"].map(_norm)
    df["gmi_n"]  = df["Gmina"].map(_norm)
    df["miej_n"] = df["Miejscowość"].map(_norm)
    df["dz_n"]   = df["Dzielnica"].map(_norm)

    return df


# ------------------- Uzupełnianie danych -------------------

def _fill_from_source(
    r: pd.Series,
    df: pd.DataFrame,
    woj_n: str,
    pow_n: str,
    gmi_n: str,
    mj_n: str,
    dz_n: str,
) -> pd.Series:
    """
    Uzupełnia braki w wierszu r na podstawie jednego źródła danych (TERYT / obszar_sadow).
    """
    if df.empty:
        return r

    subset = df

    if woj_n:
        subset = subset[subset["woj_n"] == woj_n]

    if pow_n and not subset.empty:
        tmp = subset[subset["pow_n"] == pow_n]
        if not tmp.empty:
            subset = tmp

    if dz_n and not subset.empty:
        tmp = subset[subset["dz_n"] == dz_n]
        if not tmp.empty:
            subset = tmp

    if mj_n and not subset.empty:
        tmp = subset[subset["miej_n"] == mj_n]
        if not tmp.empty:
            subset = tmp

    if gmi_n and not subset.empty:
        tmp = subset[subset["gmi_n"] == gmi_n]
        if not tmp.empty:
            subset = tmp

    if subset.empty:
        return r

    for col in ADDR_COLS:
        if _is_missing(r.get(col, "")):
            t_col = col
            if col == "Województwo":
                t_col = "Województwo" if "Województwo" in subset.columns else "Wojewodztwo"
            if col == "Miejscowość":
                t_col = "Miejscowość" if "Miejscowość" in subset.columns else "Miejscowosc"

            if t_col not in subset.columns:
                continue

            vals = subset[t_col].unique()
            if len(vals) == 1 and vals[0]:
                r[col] = vals[0]

    return r


def _enrich_row(row: pd.Series, teryt: pd.DataFrame, sad: pd.DataFrame) -> pd.Series:
    """Uzupełnia dziury adresowe w jednym wierszu."""
    r = row.copy()

    woj = r.get("Województwo", "")
    powiat = r.get("Powiat", "")
    gmi = r.get("Gmina", "")
    mj = r.get("Miejscowość", "")
    dz = r.get("Dzielnica", "")

    woj_n = _norm(woj) if not _is_missing(woj) else ""
    pow_n = _norm(powiat) if not _is_missing(powiat) else ""
    gmi_n = _norm(gmi) if not _is_missing(gmi) else ""
    mj_n  = _norm(mj) if not _is_missing(mj) else ""
    dz_n  = _norm(dz) if not _is_missing(dz) else ""

    # najpierw TERYT
    r = _fill_from_source(r, teryt, woj_n, pow_n, gmi_n, mj_n, dz_n)
    # potem obszar_sadow (jeśli coś jeszcze brakuje)
    if any(_is_missing(r.get(c, "")) for c in ADDR_COLS):
        r = _fill_from_source(r, sad, woj_n, pow_n, gmi_n, mj_n, dz_n)

    return r


# ------------------- Logika główna -------------------

def clean_report(path: Path, teryt_path: str, sad_path: str):
    if not path.exists():
        raise FileNotFoundError(f"Plik raportu nie istnieje: {path}")

    teryt = load_teryt(teryt_path)
    sad   = load_obszar_sadow(sad_path)

    df = pd.read_excel(path).fillna("")

    # upewnij się, że kolumny adresowe istnieją
    for col in ADDR_COLS:
        if col not in df.columns:
            df[col] = ""

    # zamień '---' / NaN na puste stringi
    for col in ADDR_COLS:
        df[col] = df[col].apply(lambda v: "" if _is_missing(v) else v)

    # statystyka przed
    missing_before = (df[ADDR_COLS].applymap(_is_missing)).sum()

    # uzupełnianie
    df2 = df.apply(_enrich_row, axis=1, teryt=teryt, sad=sad)

    # statystyka po
    missing_after = (df2[ADDR_COLS].applymap(_is_missing)).sum()

    # zapis nadpisujący
    df2.to_excel(path, index=False)

    print("\nCzyszczenieAdresu – statystyka braków (puste/---):")
    print("PRZED:\n", missing_before.to_string())
    print("\nPO:\n", missing_after.to_string())
    print(f"\n✔ Zapisano zmiany do pliku:\n{path}")


# ------------------- CLI -------------------

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Czyszczenie braków adresowych.")
    parser.add_argument("raport", help="Ścieżka do raportu .xlsx")
    parser.add_argument("--teryt", default="teryt.csv")
    parser.add_argument("--obszar", default="obszar_sadow.xlsx")
    args = parser.parse_args(argv)

    raport_path = Path(args.raport).resolve()

    try:
        clean_report(raport_path, teryt_path=args.teryt, sad_path=args.obszar)
    except Exception as e:
        print(f"[BŁĄD]: {e}")
        traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
