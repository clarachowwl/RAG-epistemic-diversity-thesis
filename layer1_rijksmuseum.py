"""
Layer 1 Corpus Builder — Rijksmuseum
=====================================
Fetches institutional heritage records from the Rijksmuseum
using the Linked Art API. No API key required.

Usage:
    pip install requests openpyxl
    python layer1_rijksmuseum.py
"""

import requests
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date

# ─────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────
OUTPUT_FILE = "rijksmuseum_corpus.xlsx"
MAX_PER_QUERY = 40

SEARCH_URL    = "https://data.rijksmuseum.nl/search/collection"
RESOLVER_BASE = "https://data.rijksmuseum.nl/resolve"

SEARCH_QUERIES = [
    ({"description": "Indonesia", "type": "painting"}, "Indonesia / Dutch East Indies"),
    ({"description": "Suriname slavery"},               "Suriname / Slavery"),
    ({"description": "VOC", "type": "painting"},        "VOC / Trade"),
    ({"description": "Batavia Java"},                   "Java / Batavia"),
    ({"description": "colonial portrait", "type": "painting"}, "Colonial portraiture"),
]

# ─────────────────────────────────────────
# FETCH FUNCTIONS
# ─────────────────────────────────────────
def search_identifiers(params, max_results=40):
    identifiers = []
    url = SEARCH_URL
    current_params = dict(params)
    while len(identifiers) < max_results:
        r = requests.get(url, params=current_params, timeout=15)
        if r.status_code != 200:
            break
        data = r.json()
        for item in data.get("orderedItems", []):
            identifiers.append(item["id"])
            if len(identifiers) >= max_results:
                break
        next_page = data.get("next", {}).get("id")
        if not next_page or len(identifiers) >= max_results:
            break
        url = next_page
        current_params = {}
        time.sleep(0.3)
    return identifiers


def resolve_object(linked_art_id):
    r = requests.get(linked_art_id,
                     headers={"Accept": "application/json"},
                     timeout=15,
                     allow_redirects=True)
    if r.status_code != 200:
        return None
    try:
        return r.json()
    except Exception:
        return None


def extract_fields(obj, topic_label, linked_art_id):
    if not obj:
        return None

    title = ""
    for item in obj.get("identified_by", []):
        if item.get("type") == "Name":
            title = item.get("content", "")
            break
    if not title:
        title = obj.get("_label", "")

    description = ""
    for item in obj.get("referred_to_by", []):
        content = item.get("content", "")
        lang    = item.get("language", [{}])
        lang_id = lang[0].get("id", "") if lang else ""
        if "english" in lang_id.lower() or "eng" in lang_id.lower():
            description = content
            break
        elif content and not description:
            description = content

    if not description or len(description.strip()) < 30:
        return None

    date_created = ""
    for ts in obj.get("timespan", []):
        ids = ts.get("identified_by", [{}])
        date_created = ids[0].get("content", "") if ids else str(ts.get("begin_of_the_begin", ""))[:4]
        break

    artist = "Unknown"
    for prod in obj.get("produced_by", {}).get("carried_out_by", []):
        artist = prod.get("_label", "Unknown")
        break

    obj_id = linked_art_id.split("/")[-1]

    return {
        "doc_id":              obj_id,
        "title":               title,
        "date_created":        date_created,
        "artist_name":         artist,
        "topic":               topic_label,
        "description":         description.strip(),
        "language":            "en",
        "source_institution":  "Rijksmuseum",
        "epistemic_layer":     1,
        "epistemic_stance":    "institutional",
        "source_url":          linked_art_id,
        "retrieved_date":      str(date.today()),
    }

# ─────────────────────────────────────────
# MAIN FETCH
# ─────────────────────────────────────────
def fetch_corpus():
    records  = []
    seen_ids = set()
    for params, topic_label in SEARCH_QUERIES:
        print(f"\nSearching: {params} -> '{topic_label}'")
        identifiers = search_identifiers(params, max_results=MAX_PER_QUERY)
        print(f"  Found {len(identifiers)} identifiers")
        for lid in identifiers:
            if lid in seen_ids:
                continue
            seen_ids.add(lid)
            obj    = resolve_object(lid)
            record = extract_fields(obj, topic_label, lid)
            if record:
                records.append(record)
                print(f"  + {record['doc_id']}: {record['title'][:55]}")
            else:
                print(f"  - {lid.split('/')[-1]}: no English text, skipped")
            time.sleep(0.2)
        time.sleep(0.5)
    print(f"\nTotal usable records: {len(records)}")
    return records

# ─────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────
COLUMNS = [
    ("doc_id",             "Doc ID",             18),
    ("title",              "Title",              35),
    ("date_created",       "Date Created",       14),
    ("artist_name",        "Artist Name",        24),
    ("topic",              "Topic",              26),
    ("description",        "Description (EN)",   65),
    ("language",           "Language",           10),
    ("source_institution", "Source Institution", 20),
    ("epistemic_layer",    "Epistemic Layer",    14),
    ("epistemic_stance",   "Epistemic Stance",   18),
    ("source_url",         "Source URL",         42),
    ("retrieved_date",     "Retrieved Date",     16),
]

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
ALT_FILL    = PatternFill("solid", fgColor="D6E4F0")
thin        = Side(style="thin", color="AAAAAA")
BORDER      = Border(left=thin, right=thin, top=thin, bottom=thin)


def save_to_excel(records, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rijksmuseum Corpus"
    ws.freeze_panes = "A2"

    for ci, (field, label, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(1, ci, label)
        cell.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = BORDER
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[1].height = 22

    for ri, record in enumerate(records, 2):
        alt = (ri % 2 == 0)
        for ci, (field, label, _) in enumerate(COLUMNS, 1):
            cell = ws.cell(ri, ci, record.get(field, ""))
            cell.font      = Font(size=10, name="Arial")
            cell.fill      = ALT_FILL if alt else PatternFill()
            cell.alignment = Alignment(vertical="top", wrap_text=(field == "description"))
            cell.border    = BORDER
        ws.row_dimensions[ri].height = 55 if len(str(record.get("description", ""))) > 80 else 28

    wb.save(filename)
    print(f"\nSaved: {filename} ({len(records)} records)")


if __name__ == "__main__":
    records = fetch_corpus()
    if records:
        save_to_excel(records, OUTPUT_FILE)
    else:
        print("No records fetched. Check internet connection.")
