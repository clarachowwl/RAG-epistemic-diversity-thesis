"""
Layer 3 Corpus Builder — Wikipedia Critical Sources
======================================================
Fetches decolonial and postcolonial scholarship from Wikipedia.
No API key needed.

Usage:
    pip install requests openpyxl
    python layer3_wikipedia.py
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
OUTPUT_FILE   = "layer3_critical_corpus.xlsx"
WIKIPEDIA_API = "https://en.wikipedia.org/w/api.php"

WIKIPEDIA_ARTICLES = [
    ("Dutch colonial empire",                 "Colonial history",     "historical_account"),
    ("Dutch East India Company",              "VOC / Trade",          "institutional_history"),
    ("History of Indonesia",                  "Indonesia",            "postcolonial_analysis"),
    ("Dutch East Indies",                     "Indonesia",            "institutional_history"),
    ("Batavia, Dutch East Indies",            "Java / Batavia",       "historical_account"),
    ("Java War",                              "Indonesia / Conflict", "historical_account"),
    ("Cultivation system",                    "Colonial economy",     "postcolonial_analysis"),
    ("History of Suriname",                   "Suriname",             "historical_account"),
    ("Slavery in Suriname",                   "Suriname / Slavery",   "decolonial_critique"),
    ("Atlantic slave trade",                  "Slavery",              "decolonial_critique"),
    ("Postcolonialism",                       "Theory",               "decolonial_critique"),
    ("Decolonization of knowledge",           "Theory",               "decolonial_critique"),
    ("Colonial mentality",                    "Theory",               "decolonial_critique"),
    ("Indonesian National Revolution",        "Indonesia",            "postcolonial_analysis"),
    ("Surinamese people",                     "Suriname",             "community_voice"),
    ("Javanese people",                       "Indonesia",            "community_voice"),
    ("Tropenmuseum",                          "Museum critique",      "institutional_history"),
    ("Colonial exhibition",                   "Museum critique",      "postcolonial_analysis"),
    ("Eurocentrism",                          "Theory",               "decolonial_critique"),
    ("Cultural heritage",                     "Theory",               "postcolonial_analysis"),
    ("Indigenous rights",                     "Theory",               "community_voice"),
    ("Decolonization",                        "Theory",               "decolonial_critique"),
    ("Postcolonial literature",               "Theory",               "decolonial_critique"),
]

# ─────────────────────────────────────────
# FETCH FUNCTIONS
# ─────────────────────────────────────────
def get_wikipedia_extract(title):
    params = {
        "action":    "query",
        "titles":    title,
        "prop":      "extracts|info",
        "exintro":   False,
        "explaintext": True,
        "inprop":    "url",
        "format":    "json",
        "redirects": 1,
    }
    try:
        r = requests.get(WIKIPEDIA_API, params=params, timeout=15,
                         headers={"User-Agent": "ThesisResearch/1.0"})
        r.raise_for_status()
        data  = r.json()
        pages = data.get("query", {}).get("pages", {})
        page  = next(iter(pages.values()))
        if page.get("missing") is not None:
            return None
        extract = page.get("extract", "")
        if len(extract.strip()) < 200:
            return None
        if len(extract) > 3000:
            extract = extract[:3000] + "..."
        return {
            "extract":   extract.strip(),
            "url":       page.get("fullurl", f"https://en.wikipedia.org/wiki/{title.replace(' ', '_')}"),
            "page_id":   str(page.get("pageid", "")),
            "title":     page.get("title", title),
        }
    except Exception as e:
        print(f"  Error fetching '{title}': {e}")
        return None

# ─────────────────────────────────────────
# MAIN FETCH
# ─────────────────────────────────────────
def fetch_layer3():
    records  = []
    seen_ids = set()
    print("Fetching Layer 3 — Wikipedia decolonial / critical texts\n")
    for article_title, topic, stance in WIKIPEDIA_ARTICLES:
        print(f"  -> {article_title}")
        result = get_wikipedia_extract(article_title)
        if not result:
            print(f"  Not found or too short, skipped")
            time.sleep(0.3)
            continue
        if result["page_id"] in seen_ids:
            print(f"  Duplicate, skipped")
            continue
        seen_ids.add(result["page_id"])
        records.append({
            "doc_id":                    f"wiki_{result['page_id']}",
            "title":                     result["title"],
            "author":                    "Wikipedia contributors",
            "year":                      str(date.today().year),
            "journal_publisher":         "Wikipedia (open access)",
            "doi_url":                   result["url"],
            "abstract":                  result["extract"][:500],
            "full_text_excerpt":         result["extract"],
            "keywords":                  topic,
            "geographic_focus":          topic.split("/")[0].strip(),
            "epistemic_stance":          stance,
            "author_affiliation_region": "Multiple / Collaborative",
            "language_original":         "en",
            "source_platform":           "Wikipedia",
            "epistemic_layer":           3,
            "retrieved_date":            str(date.today()),
        })
        print(f"  + {result['title']} ({len(result['extract'])} chars)")
        time.sleep(0.5)
    print(f"\nTotal Layer 3 records: {len(records)}")
    return records

# ─────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────
COLUMNS = [
    ("doc_id",                    "Doc ID",               18),
    ("title",                     "Title",                35),
    ("author",                    "Author",               22),
    ("year",                      "Year",                  8),
    ("journal_publisher",         "Journal / Publisher",  26),
    ("keywords",                  "Keywords / Topic",     24),
    ("geographic_focus",          "Geographic Focus",     18),
    ("epistemic_stance",          "Epistemic Stance",     22),
    ("abstract",                  "Abstract (500 chars)", 45),
    ("full_text_excerpt",         "Full Text Excerpt",    65),
    ("language_original",         "Language",             10),
    ("source_platform",           "Source Platform",      16),
    ("epistemic_layer",           "Epistemic Layer",      14),
    ("doi_url",                   "URL",                  42),
    ("retrieved_date",            "Retrieved Date",       16),
]

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
ALT_FILL    = PatternFill("solid", fgColor="E2EFDA")
thin        = Side(style="thin", color="AAAAAA")
BORDER      = Border(left=thin, right=thin, top=thin, bottom=thin)


def save_to_excel(records, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Layer 3 Critical Corpus"
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
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(field in ["full_text_excerpt", "abstract"]))
            cell.border    = BORDER
        ws.row_dimensions[ri].height = 60

    wb.save(filename)
    print(f"\nSaved: {filename} ({len(records)} records)")


if __name__ == "__main__":
    records = fetch_layer3()
    if records:
        save_to_excel(records, OUTPUT_FILE)
    else:
        print("No records fetched.")
