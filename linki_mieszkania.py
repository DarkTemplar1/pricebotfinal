# -*- coding: utf-8 -*-
"""
linki_mieszkania.py
Zbiera linki do ofert "SPRZEDAŻ / MIESZKANIE" z Otodom dla wskazanego województwa.

Użycie:
  python linki_mieszkania.py --region podlaskie --output podlaskie.csv
Opcjonalnie:
  --per_page 72        (domyślnie 72)
  --delay 0.60         (opóźnienie między stronami)
  --max_pages N        (na potrzeby testów)
"""

from __future__ import annotations
import argparse
import csv
import re
import sys
import time
import unicodedata
from math import ceil
from pathlib import Path
from typing import Iterable, Set, List
from urllib.parse import urljoin, urlsplit, urlunsplit

import requests
from bs4 import BeautifulSoup

UA = "Chrome/127.0.0.0"


def LOG(msg: str) -> None:
    print(msg, flush=True)


def normalize_region_slug(name: str) -> str:
    """
    Normalizacja nazwy województwa do sluga używanego przez Otodom:
      - małe litery
      - bez polskich znaków
      - spacje -> '--'
      - KAŻDY pojedynczy '-' między znakami słowa -> '--' (np. 'warminsko-mazurskie' -> 'warminsko--mazurskie')
      - tylko [a-z0-9-], bez znaków specjalnych
    """
    s = (name or "").strip().lower()
    # usuń diakrytyki
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    # spacje -> '--'
    s = re.sub(r"\s+", "--", s)
    # pojedynczy '-' pomiędzy znakami słowa -> '--'
    s = re.sub(r"(?<=\w)-(?=\w)", "--", s)
    # dopuszczalne: litery, cyfry, '-'
    s = re.sub(r"[^a-z0-9\-]+", "", s)
    # zredukuj 3+ minusy do dokładnie dwóch
    s = re.sub(r"-{3,}", "--", s)
    return s.strip("-")


def mk_session() -> requests.Session:
    s = requests.Session()
    headers = {
        "User-Agent": UA,
        "Accept-Encoding": "gzip, deflate",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Connection": "keep-alive",
        "Accept-Language": "pl-PL,pl;q=0.9,en;q=0.8",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "Referer": "https://www.otodom.pl/",
    }
    LOG(f"[HTTP] Headers: {headers}")
    s.headers.update(headers)
    return s


def soup_of(html: str) -> BeautifulSoup:
    try:
        return BeautifulSoup(html, "lxml")
    except Exception:
        return BeautifulSoup(html, "html.parser")


_BANNER_RE = re.compile(
    r"(\d+)\s*[-–]\s*(\d+)\s+og(?:ł|l)osze(?:ń|n)\s+z\s+(\d+)",
    re.IGNORECASE
)


def _int(s: str) -> int:
    return int(re.sub(r"\D", "", s)) if s else 0


def parse_banner_counts(html: str) -> tuple[int, int, int] | None:
    """
    Zwraca (lo, hi, total) z tekstu '1-72 ogłoszeń z 2798'.
    """
    m = _BANNER_RE.search(html.replace("\xa0", " "))
    if not m:
        return None
    lo = _int(m.group(1))
    hi = _int(m.group(2))
    total = _int(m.group(3))
    return lo, hi, total


def clean_url(u: str, base: str = "https://www.otodom.pl") -> str:
    """
    Normalizuje link oferty:
      - absolutny URL
      - bez query (UTM itd.)
      - zachowuje ścieżkę /pl/oferta/...
    """
    if not u:
        return ""
    absu = urljoin(base, u)
    parts = urlsplit(absu)
    # akceptuj tylko ścieżki z /pl/oferta/
    if "/pl/oferta/" not in parts.path:
        return ""
    # bez parametrów query/fragment i bez trailing slash
    return urlunsplit((parts.scheme, parts.netloc, parts.path.rstrip("/"), "", ""))


def extract_links(html: str) -> List[str]:
    """
    Główna metoda: DOM — selektor 'a[data-cy="listing-item-link"]'.
    Fallback: 'a[href*="/pl/oferta/"]'
    """
    sp = soup_of(html)
    links: list[str] = []

    # wariant podstawowy
    for a in sp.select('a[data-cy="listing-item-link"]'):
        href = a.get("href", "")
        u = clean_url(href)
        if u:
            links.append(u)

    # fallback, gdyby data-cy się zmieniło
    if not links:
        for a in sp.select('a[href*="/pl/oferta/"]'):
            href = a.get("href", "")
            u = clean_url(href)
            if u:
                links.append(u)

    return links


def unique(seq: Iterable[str]) -> List[str]:
    seen: Set[str] = set()
    out: List[str] = []
    for u in seq:
        if u and u not in seen:
            seen.add(u)
            out.append(u)
    return out


def page_url(region_slug: str, page: int, per_page: int) -> str:
    return (
        f"https://www.otodom.pl/pl/wyniki/sprzedaz/mieszkanie/"
        f"{region_slug}?limit={per_page}&ownerTypeSingleSelect=ALL&by=DEFAULT&direction=DESC&page={page}"
    )


def fetch(sess: requests.Session, url: str) -> str:
    LOG(f"[GET] {url}")
    r = sess.get(url, timeout=(10, 30), allow_redirects=True)
    LOG(f"[HTTP] status={r.status_code} final_url={r.url} len={len(r.text)}")
    r.raise_for_status()
    return r.text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--region", required=True, help="np. 'podlaskie' (bez polskich znaków też może być)")
    ap.add_argument("--output", required=True, help="ścieżka do CSV z linkami (1 kolumna: link)")
    ap.add_argument("--per_page", type=int, default=72)
    ap.add_argument("--delay", type=float, default=0.60)
    ap.add_argument("--max_pages", type=int, default=0, help="0 = wg banera; >0 ogranicza liczbę stron")
    args = ap.parse_args()

    region_input = args.region
    region_slug = normalize_region_slug(region_input)
    out_csv = Path(args.output).resolve()

    LOG(f"[start] region='{region_input}' type='mieszkanie' output='{out_csv}'")
    LOG(f"[slug] '{region_input}' -> '{region_slug}'")

    sess = mk_session()

    # Strona 1 — ustal liczbę ogłoszeń / stron
    url1 = page_url(region_slug, 1, args.per_page)
    LOG(f"[URL p1] {url1}")
    html1 = fetch(sess, url1)

    bc = parse_banner_counts(html1)
    if bc:
        lo, hi, total = bc
        LOG(f"[baner] {lo}-{hi} ogłoszeń z {total}   -> lo={lo}, hi={hi}, total={total}")
        max_pages = ceil(total / args.per_page)
        LOG(f"[pages] total={total} per_page={args.per_page} -> max_pages={max_pages}")
    else:
        LOG("[WARN] Nie udało się znaleźć banera z liczbą ogłoszeń — przyjmuję 1 stronę")
        max_pages = 1

    if args.max_pages and args.max_pages > 0:
        max_pages = min(max_pages, args.max_pages)
        LOG(f"[limit] max_pages ograniczone do {max_pages}")

    # Zbiór linków
    all_links: List[str] = []
    # Strona 1
    links1 = extract_links(html1)
    LOG(f"[page 1] dom={len(links1)} new={len(unique(links1))} total_unique={len(unique(links1))}")
    all_links.extend(links1)
    all_links = unique(all_links)
    LOG(f"[unique after p1] {len(all_links)}")

    # Następne strony
    for p in range(2, max_pages + 1):
        urlp = page_url(region_slug, p, args.per_page)
        html = fetch(sess, urlp)
        links = extract_links(html)
        before = len(all_links)
        # ile nowych?
        new_cnt = 0
        seen_set = set(all_links)
        for u in links:
            if u and u not in seen_set:
                all_links.append(u)
                seen_set.add(u)
                new_cnt += 1
        LOG(f"[page {p}] dom={len(links)} new={new_cnt} total_unique={len(all_links)}")
        if args.delay > 0:
            LOG(f"[sleep] {args.delay:.2f}s")
            time.sleep(args.delay)

    # Zapis CSV
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", encoding="utf-8-sig", newline="") as f:
        f.write("link\n")
        for u in all_links:
            f.write(u + "\n")

    LOG(f"[done] zapisano: {out_csv} (unikalnych linków: {len(all_links)})")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        LOG(f"[ERR] {e}")
        sys.exit(1)
