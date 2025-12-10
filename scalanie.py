#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
scalanie.py
Łączy wszystkie CSV z folderu "województwa" w jeden plik Excel "Polska.xlsx"
- Każde województwo trafia do osobnego arkusza
- Powstaje też arkusz zbiorczy "Polska" ze wszystkimi rekordami
Użycie:
    python scalanie.py --base "C:/Users/skalk/Desktop/yeet"
Jeśli nie podasz --base, użyje bieżącego katalogu roboczego.
"""

from __future__ import annotations
import argparse
from pathlib import Path
import pandas as pd
from datetime import datetime


def wczytaj_woj_csv(csv_path: Path) -> pd.DataFrame:
    """
    Czyta pojedynczy CSV województwa do DataFrame.
    Gwarantuje, że kolumny są zachowane tak jak w scraperze.
    Zwraca pusty DataFrame, jeśli plik jest pusty poza nagłówkiem.
    """
    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        # awaryjnie spróbuj cp1250 (czasem Excel tak zapisze)
        df = pd.read_csv(csv_path, encoding="cp1250")

    # dopisz kolumnę WOJEWODZTWO_ŹRÓDŁO jeśli nie ma (żeby w zbiorczym było wiadomo skąd rekord)
    woj_name = csv_path.stem  # np. "Podlaskie"
    if "wojewodztwo" not in df.columns:
        df["wojewodztwo"] = woj_name
    else:
        # jeżeli jest kolumna "wojewodztwo" ale jest pusta, uzupełnij
        df["wojewodztwo"] = df["wojewodztwo"].fillna(woj_name)

    return df


def scal_do_excela(base_dir: Path):
    """
    base_dir/
        województwa/
            Dolnośląskie.csv
            Mazowieckie.csv
            ...
    Tworzy base_dir/Polska.xlsx
    """
    base_dir = base_dir.resolve()
    woj_dir = base_dir / "województwa"
    out_xlsx = base_dir / "Polska.xlsx"

    if not woj_dir.exists():
        raise SystemExit(f"[ERR] Brak folderu 'województwa' w {base_dir}")

    # zbierz wszystkie csv z folderu województwa
    csv_files = sorted(woj_dir.glob("*.csv"))
    if not csv_files:
        raise SystemExit(f"[ERR] Brak plików CSV w {woj_dir}")

    arkusze: dict[str, pd.DataFrame] = {}
    all_rows = []

    for csv_path in csv_files:
        df = wczytaj_woj_csv(csv_path)

        # normalizuj nagłówki - upewnij się że wszystkie standardowe kolumny są obecne
        std_cols = [
            "cena","cena_za_metr","metry","liczba_pokoi","pietro","rynek","rok_budowy",
            "material","wojewodztwo","powiat","gmina","miejscowosc",
            "dzielnica","ulica","link"
        ]
        for col in std_cols:
            if col not in df.columns:
                df[col] = ""

        # zachowaj kolejność kolumn w arkuszach
        df = df[std_cols]

        woj_name = csv_path.stem  # nazwa arkusza np. "Dolnośląskie"
        arkusze[woj_name] = df.copy()
        all_rows.append(df.copy())

    # scal wszystko w jedną ramkę
    if all_rows:
        df_all = pd.concat(all_rows, ignore_index=True)
    else:
        df_all = pd.DataFrame()

    # zapisz do Excela:
    # - każdy województwo jako osobny sheet
    # - na końcu sheet "Polska" z całością
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        # pojedyncze województwa
        for sheet_name, df in arkusze.items():
            # Excel nie lubi bardzo długich nazw arkuszy >31 znaków, więc przytnij w razie czego
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)

        # zbiorczy
        df_all.to_excel(writer, sheet_name="Polska", index=False)

        # dodaj metadane w pustym arkuszu INFO (opcjonalnie)
        meta = pd.DataFrame(
            {
                "generated_at": [datetime.now().isoformat(timespec="seconds")],
                "source_folder": [str(woj_dir)],
                "num_regions": [len(arkusze)],
                "total_rows": [len(df_all)],
            }
        )
        meta.to_excel(writer, sheet_name="INFO", index=False)

    print(f"[OK] Zapisano plik Excel: {out_xlsx}")


def main():
    ap = argparse.ArgumentParser(
        description="Scal wszystkie województwa/*.csv do jednego Polska.xlsx"
    )
    ap.add_argument(
        "--base",
        help="Folder bazowy (tam gdzie są podfoldery 'województwa' i 'linki'). "
             "Jeśli nie podasz, użyję bieżącego katalogu.",
        default=None,
    )
    args = ap.parse_args()

    if args.base:
        base_dir = Path(args.base)
    else:
        base_dir = Path.cwd()

    scal_do_excela(base_dir)


if __name__ == "__main__":
    main()
