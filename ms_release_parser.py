#!/usr/bin/env python3
import argparse
import json
import re
import sys
from datetime import datetime

import requests
from bs4 import BeautifulSoup


DATE_FORMATS = [
    "%B %d, %Y",
    "%d %B %Y",
    "%Y-%m-%d",
]

HEADER_MAP = {
    "servicing option": "ServicingOption",
    "availability date": "AvailabilityDate",
    "build": "Build",
    "kb article": "KBArticle",
    "update type": "UpdateType",
    "type": "Type",
}

OS_PREFIXES = ("Version", "Windows Server")


# ---------- helpers ----------

def clean(text: str) -> str:
    return " ".join(text.split()).strip()


def normalize_date(text: str) -> str:
    text = text.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return text


def extract_kb(text: str):
    m = re.search(r"KB\d+", text, re.I)
    return m.group(0).upper() if m else None


def normalize_header(text: str):
    key = text.lower().strip()
    return HEADER_MAP.get(key)


def find_os_label(table):
    for prev in table.find_all_previous("strong"):
        txt = clean(prev.get_text(" ", strip=True))
        if txt.startswith(OS_PREFIXES):
            return txt
    return None


def table_has_kb(headers):
    for h in headers:
        if h and h.lower().startswith("kb"):
            return True
    return False


# ---------- parsing ----------

def parse_table(table):
    rows = table.find_all("tr")
    if not rows:
        return None

    header_cells = rows[0].find_all(["th", "td"])
    headers_raw = [clean(c.get_text(" ", strip=True)) for c in header_cells]

    if not table_has_kb(headers_raw):
        return None

    mapped = [normalize_header(h) for h in headers_raw]
    releases = []

    for tr in rows[1:]:
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue

        entry = {}

        for i, cell in enumerate(cells):
            if i >= len(mapped):
                continue

            col = mapped[i]
            if not col:
                continue

            text = clean(cell.get_text(" ", strip=True))

            if col == "AvailabilityDate":
                entry[col] = normalize_date(text)

            elif col == "KBArticle":
                kb = extract_kb(cell.get_text(" ", strip=True))
                if kb:
                    entry["KBArticle"] = kb

                # fallback: sometimes build is inside same cell
                if "Build" not in entry:
                    m = re.search(r"\d+\.\d+", text)
                    if m:
                        entry["Build"] = m.group(0)

            elif col == "Build":
                m = re.search(r"\d+\.\d+", text)
                if m:
                    entry["Build"] = m.group(0)

            else:
                entry[col] = text

        # accept rows that at least have KB + Build
        if "KBArticle" in entry and "Build" in entry:
            releases.append(entry)

    return releases if releases else None

def dedupe_releases(releases):
    """
    Remove exact duplicate dict rows.
    """
    seen = set()
    unique = []

    for r in releases:
        key = tuple(sorted(r.items()))
        if key not in seen:
            seen.add(key)
            unique.append(r)

    return unique


def sort_releases(releases):
    """
    Sort newest first by date then build.
    """
    def build_key(r):
        date = r.get("AvailabilityDate", "")
        build = r.get("Build", "")
        return (date, build)

    return sorted(releases, key=build_key, reverse=True)


def parse_release_page(url: str):
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    tables = soup.find_all("table")

    os_map = {}

    for table in tables:
        releases = parse_table(table)
        if not releases:
            continue

        os_label = find_os_label(table)
        if not os_label:
            continue

        os_map.setdefault(os_label, []).extend(releases)

    results = []

    for os_label, rels in os_map.items():
        rels = dedupe_releases(rels)
        rels = sort_releases(rels)

        results.append({
            "OS": os_label,
            "Releases": rels
        })

    # stable order of OS blocks
    results.sort(key=lambda x: x["OS"])

    return results


# ---------- main ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    try:
        data = parse_release_page(args.url)
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    main()
