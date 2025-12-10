#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
from pathlib import Path
from datetime import datetime
import traceback

def run_scraper_inside_process(in_file: Path, out_file: Path, delay_min: float, delay_max: float, retries: int, log_file):
    """
    Uruchamia scraper_otodom_mieszkania w TYM SAMYM procesie,
    przekazując parametry przez sys.argv i wywołując scraper.main().
    """
    # Import lokalny — dzięki temu PyInstaller dołączy moduł
    try:
        import scraper_otodom_mieszkania as scraper
    except Exception as e:
        print(f"[FATAL] Nie udało się zaimportować 'scraper_otodom_mieszkania': {e}", file=log_file)
        traceback.print_exc(file=log_file)
        return

    argv = [
        "scraper_otodom_mieszkania.py",
        "--input", str(in_file),
        "--output", str(out_file),
        "--delay_min", str(delay_min),
        "--delay_max", str(delay_max),
        "--retries",   str(retries),
    ]

    print("[argv]", " ".join(f'"{a}"' if " " in a else a for a in argv), file=log_file)

    old_argv = sys.argv
    try:
        sys.argv = argv
        # scraper.main() ma własny argparse i obsługę wszystkich flag
        scraper.main()
    except SystemExit as e:
        # Gdy argparse w scraperze wywoła sys.exit, złap i zapisz kod
        print(f"[i] scraper zakończył się kodem: {getattr(e, 'code', e)}", file=log_file)
    except Exception:
        print("[X] Wyjątek w scraperze:", file=log_file)
        traceback.print_exc(file=log_file)
    finally:
        sys.argv = old_argv


def main():
    ap = argparse.ArgumentParser(
        description="PriceBot Launcher — przechodzi po wszystkich plikach CSV w linki/ "
                    "i dla każdego uruchamia scraper, zapisując wynik do województwa/."
    )
    ap.add_argument(
        "--root",
        default=r"C:\Users\skalk\Desktop\baza danych",
        help="Folder bazowy zawierający podkatalogi 'linki' oraz 'województwa'."
    )
    # zachowujemy stare nazwy z minusem, ale mapujemy do sensownych dest
    ap.add_argument("--delay-min", dest="delay_min", type=float, default=4.0, help="Minimalne opóźnienie (sek).")
    ap.add_argument("--delay-max", dest="delay_max", type=float, default=6.0, help="Maksymalne opóźnienie (sek).")
    ap.add_argument("--retries", type=int, default=3, help="Liczba prób pobrania jednego ogłoszenia.")
    ap.add_argument(
        "--only",
        default="*.csv",
        help="Wzorzec nazw plików z linkami (glob). Domyślnie: *.csv"
    )
    args = ap.parse_args()

    root = Path(args.root).resolve()
    links_dir = root / "linki"
    woj_dir   = root / "województwa"
    logs_dir  = root / "logs"

    for p in (links_dir, woj_dir, logs_dir):
        p.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = logs_dir / f"launcher_{stamp}.log"

    with open(log_path, "a", encoding="utf-8") as log:
        print(f"[i] Start: {datetime.now():%Y-%m-%d %H:%M:%S}", file=log)
        print(f"[i] ROOT={root}", file=log)
        print(f"[i] LINKS_DIR={links_dir}", file=log)
        print(f"[i] WOJ_DIR={woj_dir}", file=log)
        print(f"[i] PATTERN={args.only}", file=log)
        print(f"[i] delay_min={args.delay_min}, delay_max={args.delay_max}, retries={args.retries}", file=log)

        # Lista plików z linkami
        csvs = sorted(links_dir.glob(args.only))
        if not csvs:
            print(f"[WARN] Brak plików CSV zgodnych z '{args.only}' w {links_dir}", file=log)
            print(f"[WARN] Nic do zrobienia. Log: {log_path}")
            return 0

        for in_file in csvs:
            out_file = woj_dir / in_file.name
            print("\n[run]", in_file.name, "->", out_file.name, file=log)

            # Uruchom scraper w tym samym procesie (bez subprocess)
            run_scraper_inside_process(
                in_file=in_file,
                out_file=out_file,
                delay_min=args.delay_min,
                delay_max=args.delay_max,
                retries=args.retries,
                log_file=log
            )

        print(f"\n[OK] Zakończono: {datetime.now():%Y-%m-%d %H:%M:%S}", file=log)

    print(f"Gotowe. Log: {log_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
