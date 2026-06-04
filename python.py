"""
Company Name Normalizer  (Nexl = source of truth)
-------------------------------------------------
Policy: every survey company gets the BEST Nexl match written into the verified
column with its score. Only names with NO usable candidate are marked
"REVIEW MANUALLY".

How to run:
    pip install pandas openpyxl rapidfuzz
    python company_name_normalizer.py
"""

import re
import pandas as pd
from collections import defaultdict
from rapidfuzz import process, fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ============================================================
# CONFIG  — edit paths / columns for your files
# ============================================================
NEXL_CSV   = "Companies-1780421663.csv"
SURVEY_IN  = "test_case.xlsx"
SURVEY_OUT = SURVEY_IN              # write results back into the same file

# (sheet name, company column, verified-name column, accuracy column) — 1-indexed
SHEETS = [
    ("Fortune 500 + Corp. Connect",     2, 8, 9),
    ("AM Law 200 + Legal Connectivity", 2, 5, 6),
]

# A candidate below this score is treated as "no real match" -> REVIEW MANUALLY.
MIN_SCORE = 70

# Targets that are generic fragments: even at a high score these are almost
# always a wrong collision (e.g. "United Parcel Service" -> "United"), so we
# send them to review instead of writing a misleading match.
DENY_TARGETS = {
    "law office", "united", "the independent", "real", "southern company",
    "cleveland.com", "post holdings", "the guardian", "marcus corporation",
    "the wendy's company", "west corporation", "fidelity international",
    "peter", "bell", "knight",
}

# ============================================================
# NORMALIZATION
# ============================================================
FORM_SUFFIXES = {
    "llp", "llc", "inc", "ltd", "limited", "plc", "corp", "corporation", "co",
    "company", "companies", "pc", "lp", "lllp", "gmbh", "ag", "sa", "nv", "pllc", "pa",
}
DESCRIPTIVE = {"group", "holdings", "holding", "intl", "international"}
STOPWORDS   = {"and", "of", "the", "for", "a", "an"}
GENERIC     = STOPWORDS | FORM_SUFFIXES | DESCRIPTIVE


def normalize(name):
    """Lowercase, drop punctuation/urls/'&', strip a leading 'the' and
    corporate-form suffixes, collapse whitespace."""
    if pd.isna(name):
        return ""
    n = str(name).lower().strip()
    n = n.replace("&", " and ")
    n = re.sub(r"https?://", " ", n)
    n = re.sub(r"\b(www\.|\.com|\.net|\.org|\.io|\.ai)\b", " ", n)
    n = re.sub(r"[^a-z0-9 ]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    if n.startswith("the "):
        n = n[4:]
    toks = [t for t in n.split() if t not in FORM_SUFFIXES]
    return " ".join(toks).strip() if toks else n


def core_tokens(clean):
    toks = [t for t in clean.split() if t not in GENERIC]
    return toks if toks else clean.split()


def _domainish(raw):
    return bool(re.search(r"\.(com|io|net|org|ai|co|us)\b", str(raw).lower()))


# ============================================================
# BUILD CANONICAL STRUCTURES FROM NEXL
# ============================================================
nexl = pd.read_csv(NEXL_CSV, low_memory=False)
nexl["clean"] = nexl["Name"].apply(normalize)

clean_to_names = defaultdict(list)
for raw, c in zip(nexl["Name"], nexl["clean"]):
    if c:
        clean_to_names[c].append(str(raw))

choices    = list(clean_to_names.keys())
rep_name   = {c: sorted(set(v), key=lambda x: (len(x), x))[0] for c, v in clean_to_names.items()}
clean_core = {c: set(core_tokens(c)) for c in choices}


# ============================================================
# MATCHER  — returns (best_name | None, score)
# ============================================================
def best_match(raw):
    c = normalize(raw)
    if not c:
        return None, 0

    # 1) exact cleaned match
    if c in clean_to_names:
        return sorted(set(clean_to_names[c]))[0], 100

    survey_core = set(core_tokens(c))
    survey_list = core_tokens(c)

    # 2) survey name fully contained in exactly one fuller official name
    contained = [
        k for k in choices
        if survey_core and survey_core <= clean_core[k]
        and len(survey_core) / max(len(clean_core[k]), 1) >= 0.5
    ]
    if len(contained) == 1:
        nm = rep_name[contained[0]]
        return nm, round(fuzz.WRatio(c, normalize(nm)), 1)

    # 3) reverse containment: a distinctive Nexl name is the leading part of the
    #    survey name (e.g. "Verizon Communications" -> "Verizon")
    rc = []
    for k in choices:
        kc = clean_core[k]
        if kc and kc < survey_core and any(len(t) >= 4 for t in kc):
            ktok = [t for t in survey_list if t in kc]
            if survey_list[:len(ktok)] == ktok and not _domainish(rep_name[k]):
                rc.append(k)
    if len(rc) == 1:
        nm = rep_name[rc[0]]
        return nm, round(fuzz.WRatio(c, normalize(nm)), 1)

    # 4) fuzzy best candidate
    res = process.extract(c, choices, scorer=fuzz.WRatio, limit=1, score_cutoff=60)
    if res:
        best_clean = res[0][0]
        score = min(res[0][1], fuzz.token_sort_ratio(c, best_clean) + 8)
        return rep_name[best_clean], round(score, 1)

    return None, 0


# ============================================================
# PROCESS WORKBOOK
# ============================================================
red_fill = PatternFill(fill_type="solid", start_color="FFB3B3", end_color="FFB3B3")
wb = load_workbook(SURVEY_IN, data_only=False)   # keeps existing formulas intact

review_rows = []
summary = {}


def process_sheet(sheet_name, company_col, verified_col, accuracy_col):
    ws = wb[sheet_name]
    total = matched = review = 0

    for r in range(2, ws.max_row + 1):
        company = ws.cell(r, company_col).value
        if not company:
            continue
        total += 1

        name, score = best_match(company)
        v_cell = ws.cell(r, verified_col)
        a_cell = ws.cell(r, accuracy_col)

        no_candidate = (
            name is None
            or (isinstance(score, (int, float)) and score < MIN_SCORE)
            or normalize(name) in DENY_TARGETS
        )

        if no_candidate:
            v_cell.value = "REVIEW MANUALLY"
            a_cell.value = None                       # no misleading number
            v_cell.fill = a_cell.fill = red_fill
            review += 1
            review_rows.append([sheet_name, r, str(company).strip()])
        else:
            v_cell.value = name
            a_cell.value = score
            matched += 1

    summary[sheet_name] = (total, matched, review)
    return total, matched, review


for sheet_name, c_col, v_col, a_col in SHEETS:
    process_sheet(sheet_name, c_col, v_col, a_col)

# small sheet listing only the rows that need a human
if "Match Review" in wb.sheetnames:
    del wb["Match Review"]
rv = wb.create_sheet("Match Review")
rv.append(["Sheet", "Row", "Survey Company (no usable Nexl match)"])
for row in review_rows:
    rv.append(row)

wb.save(SURVEY_OUT)

# ============================================================
# SUMMARY
# ============================================================
print("\n========== SUMMARY ==========")
T = M = R = 0
for sheet_name, (t, m, r) in summary.items():
    print(f"\n{sheet_name}")
    print(f"  Total: {t}   Matched (name + score): {m}   Review manually: {r}")
    T += t; M += m; R += r
print("\nOVERALL")
print(f"  Total: {T}   Matched: {M}   Review manually: {R}")
print(f"\nSaved: {SURVEY_OUT}")