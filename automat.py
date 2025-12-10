#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

"""
automat.py - wersja z:
- GUI progów ludności,
- PopulationResolver + cache + (opcjonalnie) BDL GUS API,
- korektą: Średnia skorygowana cena za m2 = Średnia cena za m2 ( z bazy) * (100% - % negocjacyjny).

Uruchamiany z selektor_csv.py jako:

    automat.main(["automat.py", RAPORT_PATH, BAZA_FOLDER])
"""

from pathlib import Path
import sys
import unicodedata
import csv
from typing import Optional, Dict, Tuple

import os
import datetime

import pandas as pd
import numpy as np

import tkinter as tk
from tkinter import ttk

import requests  # do BDL GUS API


# ---------- helpers ----------

def _norm(s: str) -> str:
    return (s or "").strip().lower().replace(" ", "").replace("\xa0", "").replace("\t", "")


def _plain(s: str) -> str:
    s = (s or "").lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return s


def _find_col(cols, candidates):
    norm_map = {_norm(c): c for c in cols}
    # najpierw pełne dopasowanie
    for cand in candidates:
        key = _norm(cand)
        if key in norm_map:
            return norm_map[key]
    # potem "zawiera"
    for c in cols:
        if any(_norm(x) in _norm(c) for x in candidates):
            return c
    return None


def _trim_after_semicolon(val):
    if pd.isna(val):
        return ""
    s = str(val)
    if ";" in s:
        s = s.split(";", 1)[0]
    return s.strip()


def _to_float_maybe(x):
    """Parsuje liczby typu '101,62 m²', '52 m2', '11 999 zł/m²' itd."""
    if pd.isna(x):
        return None
    s = str(x)
    for unit in ["m²", "m2", "zł/m²", "zł/m2", "zł"]:
        s = s.replace(unit, "")
    s = s.replace(" ", "").replace("\xa0", "")
    s = s.replace(",", ".")
    try:
        return float(s) if s else None
    except Exception:
        return None


VALUE_COLS = [
    "Średnia cena za m2 ( z bazy)",
    "Średnia skorygowana cena za m2",
    "Statystyczna wartość nieruchomości",
]

# ---------- PROGI LUDNOŚCI + DOMYŚLNE USTAWIENIA ----------

# Domyślne wartości jak na screenie:
# 0–20 000      → 25 m², 15%
# 20 000–50 000 → 20 m², 15%
# 50 000–200 000→ 15 m², 15%
# 200 000+      → 10 m², 15%
POP_MARGIN_RULES = [
    (0,      20000,   25.0, 15.0),
    (20000,  50000,   20.0, 15.0),
    (50000,  200000,  15.0, 15.0),
    (200000, None,    10.0, 15.0),
]


# ---------- KONFIGURACJA BDL (GUS) ----------

BDL_BASE_URL = "https://bdl.stat.gov.pl/api/v1"
BDL_POP_SUBJECT_ID = "P2425"  # temat „ludność” w BDL (do ewentualnego doprecyzowania)

_BDL_POP_VAR_ID: str | None = None  # id zmiennej „ludność ogółem” (cache w RAM)


def configure_margins_gui():
    """
    Okno GUI z progami ludności.

    Kolumny "minimalna" i "maksymalna ludność" są tylko do odczytu.
    Użytkownik może zmieniać:
    - "Pomiar brzegowy m²"
    - "% negocjacyjny"

    Zwraca nową listę [(low, high, m2, pct), ...] albo None (Anuluj).
    """
    root = tk.Tk()
    root.title("Ustawienia progów ludności")
    root.resizable(False, False)

    ttk.Label(root, text="Minimalna ludność").grid(row=0, column=0, padx=4, pady=4)
    ttk.Label(root, text="Maksymalna ludność").grid(row=0, column=1, padx=4, pady=4)
    ttk.Label(root, text="Pomiar brzegowy m²").grid(row=0, column=2, padx=4, pady=4)
    ttk.Label(root, text="% negocjacyjny").grid(row=0, column=3, padx=4, pady=4)

    entries_m2: list[ttk.Entry] = []
    entries_pct: list[ttk.Entry] = []

    def _fmt_pop(x):
        if x is None:
            return "∞"
        try:
            x = int(x)
        except Exception:
            return str(x)
        return f"{x:,}".replace(",", " ")

    for i, (low, high, m2, pct) in enumerate(POP_MARGIN_RULES, start=1):
        ttk.Label(root, text=_fmt_pop(low)).grid(
            row=i, column=0, padx=4, pady=2, sticky="e"
        )
        ttk.Label(root, text=_fmt_pop(high)).grid(
            row=i, column=1, padx=4, pady=2, sticky="e"
        )

        e_m2 = ttk.Entry(root, width=8, justify="right")
        e_m2.insert(0, str(m2))
        e_m2.grid(row=i, column=2, padx=4, pady=2)
        entries_m2.append(e_m2)

        e_pct = ttk.Entry(root, width=8, justify="right")
        e_pct.insert(0, str(pct))
        e_pct.grid(row=i, column=3, padx=4, pady=2)
        entries_pct.append(e_pct)

    result = {"ok": False, "rules": POP_MARGIN_RULES}

    def on_ok():
        new_rules = []
        for (low, high, default_m2, default_pct), e_m2, e_pct in zip(
            POP_MARGIN_RULES, entries_m2, entries_pct
        ):
            raw_m2 = e_m2.get().strip().replace(" ", "").replace(",", ".")
            raw_pct = e_pct.get().strip().replace(" ", "").replace(",", ".")
            try:
                m2_val = float(raw_m2) if raw_m2 else float(default_m2)
            except Exception:
                m2_val = float(default_m2)
            try:
                pct_val = float(raw_pct) if raw_pct else float(default_pct)
            except Exception:
                pct_val = float(default_pct)
            new_rules.append((low, high, m2_val, pct_val))
        result["ok"] = True
        result["rules"] = new_rules
        root.destroy()

    def on_cancel():
        result["ok"] = False
        root.destroy()

    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=len(POP_MARGIN_RULES) + 1, column=0, columnspan=4, pady=(8, 8))

    ttk.Button(btn_frame, text="Anuluj", command=on_cancel).pack(side="right", padx=4)
    ttk.Button(btn_frame, text="Start", command=on_ok).pack(side="right", padx=4)

    # wyśrodkowanie okna
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

    root.mainloop()

    if not result["ok"]:
        return None
    return result["rules"]


def rules_for_population(pop):
    """Zwraca (margines_m², %_negocjacyjny) na podstawie liczby mieszkańców."""
    if pop is None:
        return float(POP_MARGIN_RULES[-1][2]), float(POP_MARGIN_RULES[-1][3])
    try:
        p = float(pop)
    except Exception:
        return float(POP_MARGIN_RULES[-1][2]), float(POP_MARGIN_RULES[-1][3])

    for low, high, m2, pct in POP_MARGIN_RULES:
        if p >= low and (high is None or p < high):
            return float(m2), float(pct)
    return float(POP_MARGIN_RULES[-1][2]), float(POP_MARGIN_RULES[-1][3])


def _eq_mask(df: pd.DataFrame, col_candidates, value: str) -> pd.Series:
    col = _find_col(df.columns, col_candidates)
    if col is None or not str(value).strip():
        return pd.Series(True, index=df.index)
    s = df[col].astype(str).str.strip().str.lower()
    v = str(value).strip().lower()
    return s == v


# ---------- PopulationResolver ----------

class PopulationResolver:
    """
    Klasa pomocnicza do ustalania liczby mieszkańców miejscowości.

    - najpierw sprawdza lokalny cache CSV (population_cache.csv),
    - jeśli brak wpisu i włączony use_api=True, wywołuje zewnętrzne API (GUS BDL),
    - zapisuje nowe wpisy do cache.
    """

    def __init__(self, cache_path: Path, use_api: bool = False):
        self.cache_path = cache_path
        self.use_api = use_api
        self._cache: Dict[str, Tuple[float, str, str, str, str]] = {}
        self._dirty = False
        self._load_cache()

    def _make_key(self, woj: str, powiat: str, gmina: str, miejscowosc: str) -> str:
        parts = [woj or "", powiat or "", gmina or "", miejscowosc or ""]
        parts = [_plain(p) for p in parts]
        return "|".join(parts)

    def _load_cache(self):
        if not self.cache_path or not self.cache_path.exists():
            return
        try:
            with self.cache_path.open("r", encoding="utf-8-sig", newline="") as f:
                rd = csv.DictReader(f)
                for row in rd:
                    try:
                        pop = float(row.get("population", "") or 0)
                    except ValueError:
                        continue
                    key = row.get("key") or self._make_key(
                        row.get("woj", ""), row.get("powiat", ""),
                        row.get("gmina", ""), row.get("miejscowosc", "")
                    )
                    self._cache[key] = (
                        pop,
                        row.get("woj", ""),
                        row.get("powiat", ""),
                        row.get("gmina", ""),
                        row.get("miejscowosc", ""),
                    )
        except Exception as e:
            print(f"[PopulationResolver] Nie udało się wczytać cache: {e}")

    def _save_cache(self):
        if not self._dirty or not self.cache_path:
            return
        try:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            with self.cache_path.open("w", encoding="utf-8-sig", newline="") as f:
                fieldnames = ["key", "woj", "powiat", "gmina", "miejscowosc", "population"]
                wr = csv.DictWriter(f, fieldnames=fieldnames)
                wr.writeheader()
                for key, (pop, woj, pow, gmi, mia) in self._cache.items():
                    wr.writerow(
                        {
                            "key": key,
                            "woj": woj,
                            "powiat": pow,
                            "gmina": gmi,
                            "miejscowosc": mia,
                            "population": pop,
                        }
                    )
            self._dirty = False
        except Exception as e:
            print(f"[PopulationResolver] Nie udało się zapisać cache: {e}")

    def _get_bdl_headers(self) -> dict:
        """
        Zwraca nagłówki HTTP do API BDL z kluczem API.
        Klucz pobierany jest z:
        - zmiennej środowiskowej BDL_API_KEY lub GUS_BDL_API_KEY
        """
        api_key = (
            os.getenv("BDL_API_KEY")
            or os.getenv("GUS_BDL_API_KEY")
        )
        if not api_key:
            print(
                "[PopulationResolver] Brak klucza API BDL. "
                "Ustaw zmienną środowiskową BDL_API_KEY."
            )
            return {}
        return {
            "X-ClientId": api_key,
            "Accept": "application/json",
        }

    def _get_population_var_id(self) -> str | None:
        """
        Pobiera (i zapamiętuje) ID zmiennej 'ludność ogółem' z BDL.

        Szuka w grupie tematycznej P2425 (ludność).
        Jeśli nie znajdzie automatycznie – można wpisać ID ręcznie w _BDL_POP_VAR_ID.
        """
        global _BDL_POP_VAR_ID
        if _BDL_POP_VAR_ID:
            return _BDL_POP_VAR_ID

        headers = self._get_bdl_headers()
        if not headers:
            return None

        try:
            url = f"{BDL_BASE_URL}/variables"
            params = {
                "subject-id": BDL_POP_SUBJECT_ID,
                "page-size": 100,
                "format": "json",
            }
            r = requests.get(url, headers=headers, params=params, timeout=10)
            r.raise_for_status()
            data = r.json()
            for v in data.get("results", []):
                name = (v.get("name") or v.get("n1") or "").lower()
                if "ludność ogółem" in name or "ludnosc ogolem" in name or "population total" in name:
                    _BDL_POP_VAR_ID = str(v["id"])
                    print(f"[PopulationResolver] Zmienna ludności: id={_BDL_POP_VAR_ID} ({name})")
                    return _BDL_POP_VAR_ID

            print(
                "[PopulationResolver] Nie znalazłem zmiennej 'ludność ogółem' automatycznie. "
                "Możesz wpisać ręcznie ID zmiennej w _BDL_POP_VAR_ID."
            )
            return None
        except Exception as e:
            print(f"[PopulationResolver] Błąd przy pobieraniu listy zmiennych BDL: {e}")
            return None

    def _fetch_population_from_api(
        self, woj: str, powiat: str, gmina: str, miejscowosc: str
    ) -> Optional[float]:
        """
        Próbuje pobrać liczbę mieszkańców z API BDL GUS.

        Uproszczony algorytm:
        1) wyszukuje jednostkę terytorialną po nazwie miejscowości/gminy,
        2) dla znalezionego id jednostki pobiera wartość zmiennej 'ludność ogółem'
           za ostatni dostępny rok.
        """
        headers = self._get_bdl_headers()
        if not headers:
            return None

        name_search = miejscowosc or gmina
        if not name_search:
            return None

        # 1) jednostka terytorialna (przyjmijmy level=6 – gminy / miasta)
        try:
            url_units = f"{BDL_BASE_URL}/units"
            params_units = {
                "name": name_search,
                "level": "6",
                "page-size": 50,
                "format": "json",
            }
            ru = requests.get(url_units, headers=headers, params=params_units, timeout=10)
            if ru.status_code != 200:
                print(f"[PopulationResolver] BDL /units error {ru.status_code}: {ru.text[:200]}")
                return None
            ju = ru.json()
            units = ju.get("results", [])
            if not units:
                print(f"[PopulationResolver] Brak jednostek BDL dla nazwy '{name_search}'.")
                return None

            def _score(u):
                nm = (u.get("name") or "").lower()
                sc = 0
                if _plain(name_search) in _plain(nm):
                    sc += 2
                if woj and _plain(woj) in _plain(nm):
                    sc += 1
                if powiat and _plain(powiat) in _plain(nm):
                    sc += 1
                return sc

            units.sort(key=_score, reverse=True)
            unit_id = units[0].get("id")
            if not unit_id:
                print("[PopulationResolver] Nie udało się ustalić id jednostki BDL.")
                return None
        except Exception as e:
            print(f"[PopulationResolver] Błąd przy wyszukiwaniu jednostki w BDL: {e}")
            return None

        # 2) id zmiennej 'ludność ogółem'
        var_id = self._get_population_var_id()
        if not var_id:
            return None

        year = datetime.date.today().year - 1

        try:
            url_data = f"{BDL_BASE_URL}/data/by-unit/{unit_id}"
            params_data = {
                "var-id": var_id,
                "year": str(year),
                "format": "json",
            }
            rd = requests.get(url_data, headers=headers, params=params_data, timeout=10)
            if rd.status_code != 200:
                print(f"[PopulationResolver] BDL /data/by-unit error {rd.status_code}: {rd.text[:200]}")
                return None

            jd = rd.json()
            results = jd.get("results") or jd.get("series") or []
            if not results:
                print(f"[PopulationResolver] Brak wyników BDL dla jednostki {unit_id}, rok {year}.")
                return None

            entry = results[0]
            vals = entry.get("values") or entry.get("value") or []
            if isinstance(vals, list):
                for v in vals:
                    if isinstance(v, dict):
                        raw = v.get("val")
                    else:
                        raw = v
                    if raw not in (None, ""):
                        try:
                            pop = float(str(raw).replace(" ", "").replace(",", "."))
                            return pop
                        except Exception:
                            continue
            else:
                try:
                    pop = float(str(vals).replace(" ", "").replace(",", "."))
                    return pop
                except Exception:
                    pass

            return None
        except Exception as e:
            print(f"[PopulationResolver] Błąd przy pobieraniu danych BDL: {e}")
            return None

    def get_population(self, woj: str, powiat: str, gmina: str, miejscowosc: str) -> Optional[float]:
        key = self._make_key(woj, powiat, gmina, miejscowosc)
        if key in self._cache:
            return self._cache[key][0]

        pop = None
        if self.use_api:
            pop = self._fetch_population_from_api(woj, powiat, gmina, miejscowosc)

        if pop is None:
            return None

        self._cache[key] = (float(pop), woj or "", powiat or "", gmina or "", miejscowosc or "")
        self._dirty = True
        self._save_cache()
        return float(pop)


# ---------- Przetwarzanie jednego wiersza ----------

def _process_row(
    df_raport: pd.DataFrame,
    idx: int,
    df_pl: pd.DataFrame,
    col_area_pl: str,
    col_price_pl: str,
    margin_m2_default: float,
    margin_pct_default: float,
    pop_resolver: PopulationResolver,
) -> None:

    row = df_raport.iloc[idx]

    kw_col = _find_col(
        df_raport.columns,
        ["Nr KW", "nr_kw", "nrksiegi", "nr księgi", "nr_ksiegi", "numer księgi"],
    )
    kw_value = (str(row[kw_col]).strip()
                if (kw_col and pd.notna(row[kw_col]) and str(row[kw_col]).strip())
                else f"WIERSZ_{idx+1}")

    area_col = _find_col(df_raport.columns, ["Obszar", "metry", "powierzchnia"])
    area_val = _to_float_maybe(_trim_after_semicolon(row[area_col])) if area_col else None
    if area_val is None:
        print(f"[Automat] Wiersz {idx+1}: brak obszaru – pomijam.")
        return

    def _get(cands):
        c = _find_col(df_raport.columns, cands)
        return _trim_after_semicolon(row[c]) if c else ""

    woj_r = _get(["Województwo", "wojewodztwo", "woj"])
    pow_r = _get(["Powiat"])
    gmi_r = _get(["Gmina"])
    mia_r = _get(["Miejscowość", "Miasto", "miejscowosc", "miasto"])
    dzl_r = _get(["Dzielnica", "Osiedle"])
    uli_r = _get(["Ulica", "Ulica(dla budynku)", "Ulica(dla lokalu)"])

    # 1. próba użycia kolumny z liczbą mieszkańców w raporcie
    pop_col = _find_col(
        df_raport.columns,
        [
            "Liczba mieszkańców",
            "Liczba mieszkancow",
            "Wielkosc miejscowosci",
            "Wielkość miejscowości",
        ],
    )
    pop_val = None
    if pop_col:
        pop_val = _to_float_maybe(_trim_after_semicolon(row[pop_col]))

    # 2. jeśli brak – PopulationResolver (cache / API)
    if pop_val is None and pop_resolver is not None:
        pop_val = pop_resolver.get_population(woj_r, pow_r, gmi_r, mia_r)

    # 3. margines m² + % negocjacyjny (z progów)
    if pop_val is not None:
        margin_m2_row, margin_pct_row = rules_for_population(pop_val)
        print(
            f"[Automat] {kw_value}: miejscowość '{mia_r}' (pop={pop_val}) → "
            f"margines {margin_m2_row} m², % negocjacyjny {margin_pct_row}."
        )
    else:
        margin_m2_row = float(margin_m2_default or 0.0)
        margin_pct_row = float(margin_pct_default or 0.0)
        print(
            f"[Automat] {kw_value}: brak danych o liczbie mieszkańców – "
            f"używam marginesu globalnego {margin_m2_row} m² oraz % negocjacyjnego {margin_pct_row}."
        )

    # efektywny % negocjacyjny – obniżka: 100% - %negocjacyjny
    margin_pct_effective = float(margin_pct_row or 0.0)

    # zakres metrażu
    delta = abs(margin_m2_row)
    low, high = max(0.0, area_val - delta), area_val + delta

    m = df_pl[col_area_pl].map(_to_float_maybe)
    mask_area = (m >= low) & (m <= high)

    mask_full = mask_area.copy()
    mask_full &= _eq_mask(df_pl, ["wojewodztwo", "województwo"], woj_r)
    mask_full &= _eq_mask(df_pl, ["powiat"], pow_r)
    mask_full &= _eq_mask(df_pl, ["gmina"], gmi_r)
    mask_full &= _eq_mask(df_pl, ["miejscowosc", "miasto", "miejscowość"], mia_r)
    if dzl_r:
        mask_full &= _eq_mask(df_pl, ["dzielnica", "osiedle"], dzl_r)
    if uli_r:
        mask_full &= _eq_mask(df_pl, ["ulica"], uli_r)

    df_sel = df_pl[mask_full].copy()

    # fallbacky lokalizacyjne
    if df_sel.empty and uli_r:
        mask_ul = mask_area.copy()
        mask_ul &= _eq_mask(df_pl, ["wojewodztwo", "województwo"], woj_r)
        mask_ul &= _eq_mask(df_pl, ["miejscowosc", "miasto", "miejscowość"], mia_r)
        if dzl_r:
            mask_ul &= _eq_mask(df_pl, ["dzielnica", "osiedle"], dzl_r)
        mask_ul &= _eq_mask(df_pl, ["ulica"], uli_r)
        df_sel = df_pl[mask_ul].copy()

    if df_sel.empty and dzl_r:
        mask_dziel = mask_area.copy()
        mask_dziel &= _eq_mask(df_pl, ["wojewodztwo", "województwo"], woj_r)
        mask_dziel &= _eq_mask(df_pl, ["miejscowosc", "miasto", "miejscowość"], mia_r)
        mask_dziel &= _eq_mask(df_pl, ["dzielnica", "osiedle"], dzl_r)
        df_sel = df_pl[mask_dziel].copy()

    if df_sel.empty:
        mask_city = mask_area.copy()
        mask_city &= _eq_mask(df_pl, ["wojewodztwo", "województwo"], woj_r)
        mask_city &= _eq_mask(df_pl, ["miejscowosc", "miasto", "miejscowość"], mia_r)
        df_sel = df_pl[mask_city].copy()

    if df_sel.empty:
        print(f"[Automat] {kw_value}: brak dopasowanych rekordów w bazie (po filtrach).")
        return

    prices = df_sel[col_price_pl].map(_to_float_maybe).dropna()
    if prices.empty:
        print(f"[Automat] {kw_value}: brak cen w dopasowanych rekordach.")
        return

    mean_price = prices.mean()

    mean_col = _find_col(df_raport.columns, ["Średnia cena za m2 ( z bazy)", "Srednia cena za m2 ( z bazy)"])
    corr_col = _find_col(df_raport.columns, ["Średnia skorygowana cena za m2", "Srednia skorygowana cena za m2"])
    val_col = _find_col(df_raport.columns, ["Statystyczna wartość nieruchomości", "Statystyczna wartosc nieruchomosci"])

    if not mean_col or not corr_col or not val_col:
        print(f"[Automat] {kw_value}: brak wymaganych kolumn wynikowych w raporcie.")
        return

    # 1) wpisujemy średnią do „Średnia cena za m2 ( z bazy)”
    df_raport.at[idx, mean_col] = mean_price

    # 2) Średnia skorygowana cena za m2 =
    #    Średnia cena za m2 ( z bazy) * (100% − % negocjacyjny)
    baza_val = df_raport.at[idx, mean_col]
    try:
        baza_val_f = float(baza_val)
    except Exception:
        baza_val_f = float(mean_price)

    corrected_price = baza_val_f * (1.0 - margin_pct_effective / 100.0)
    df_raport.at[idx, corr_col] = corrected_price

    # 3) Statystyczna wartość nieruchomości = skorygowana cena * metry
    value = corrected_price * area_val
    df_raport.at[idx, val_col] = value

    print(
        f"[Automat] {kw_value}: dopasowano {len(df_sel)} rekordów, "
        f"średnia cena {mean_price:.2f}, "
        f"skorygowana (po negocjacji {margin_pct_effective:.1f}%) {corrected_price:.2f}, "
        f"wartość {value:.2f}."
    )


# ---------- MAIN ----------

def main(argv=None) -> int:
    global POP_MARGIN_RULES  # ważne: deklaracja global na początku funkcji

    if argv is None:
        argv = sys.argv

    if len(argv) < 3:
        print("Użycie: automat.py RAPORT_PATH BAZA_FOLDER")
        return 1

    raport_path = Path(argv[1]).resolve()
    baza_folder = Path(argv[2]).resolve()

    if not raport_path.exists():
        print(f"[BŁĄD] Nie znaleziono raportu: {raport_path}")
        return 1

    polska_path = baza_folder / "Polska.xlsx"
    if not polska_path.exists():
        print(f"[BŁĄD] Nie znaleziono Polska.xlsx w folderze: {baza_folder}")
        return 1

    # domyślne globalne – zostaną zaraz nadpisane na podstawie GUI
    margin_m2_default = 15.0
    margin_pct_default = 15.0

    # --- okno ustawień progów ludności ---
    try:
        new_rules = configure_margins_gui()
    except Exception as e:
        print(f"[Automat] Błąd GUI progów ludności: {e}")
        new_rules = POP_MARGIN_RULES

    if new_rules is None:
        print("[Automat] Przerwano działanie (Anuluj w oknie progów ludności).")
        return 1

    # zaktualizuj globalne progi
    POP_MARGIN_RULES = new_rules

    # przykładowo: jako globalny margines weź trzeci próg (50–200k)
    try:
        if len(POP_MARGIN_RULES) >= 3:
            margin_m2_default = float(POP_MARGIN_RULES[2][2])
            margin_pct_default = float(POP_MARGIN_RULES[2][3])
    except Exception:
        pass

    # wczytaj bazę
    try:
        df_pl = pd.read_excel(polska_path)
    except Exception as e:
        print(f"[BŁĄD] Nie mogę wczytać Polska.xlsx: {polska_path}\n{e}")
        return 1

    col_area_pl = _find_col(df_pl.columns, ["metry", "powierzchnia", "Obszar", "obszar"])
    col_price_pl = _find_col(df_pl.columns, ["cena_za_metr", "cena za metr", "cena_za_m2", "cena_za_metr2", "cena za m2"])
    if not col_area_pl or not col_price_pl:
        print("[BŁĄD] Polska.xlsx nie zawiera wymaganych kolumn metrażu / ceny.")
        return 1

    # wczytaj raport
    try:
        if raport_path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
            df_raport = pd.read_excel(raport_path)
            is_excel = True
        else:
            df_raport = pd.read_csv(raport_path, sep=None, engine="python")
            is_excel = False
    except Exception as e:
        print(f"[BŁĄD] Nie mogę wczytać raportu: {raport_path}\n{e}")
        return 1

    # inicjalizacja PopulationResolver z API włączonym
    pop_cache_path = baza_folder / "population_cache.csv"
    pop_resolver = PopulationResolver(pop_cache_path, use_api=True)

    n_rows = len(df_raport.index)
    print(f"[Automat] Start – liczba wierszy w raporcie: {n_rows}")

    for idx in range(n_rows):
        try:
            _process_row(
                df_raport=df_raport,
                idx=idx,
                df_pl=df_pl,
                col_area_pl=col_area_pl,
                col_price_pl=col_price_pl,
                margin_m2_default=margin_m2_default,
                margin_pct_default=margin_pct_default,
                pop_resolver=pop_resolver,
            )
        except Exception as e:
            print(f"[Automat] Błąd przy wierszu {idx+1}: {e}")

    # zapis
    try:
        if is_excel:
            df_raport.to_excel(raport_path, index=False)
        else:
            df_raport.to_csv(raport_path, index=False, encoding="utf-8-sig")
    except Exception as e:
        print(f"[BŁĄD] Nie udało się zapisać raportu: {raport_path}\n{e}")
        return 1

    print(f"[Automat] Zakończono – zapisano zmiany w pliku: {raport_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
