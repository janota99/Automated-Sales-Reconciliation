import pandas as pd
import re
from difflib import get_close_matches


def clean_alphanumeric(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(val).strip().upper())


def clean_po(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    for noise in ["P.O.", "P.O", "PO#", "PO-", "PO"]:
        s = s.replace(noise, "")
    return clean_alphanumeric(s)


def normalize_amount(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    try:
        s = str(val).replace("$", "").replace(",", "").strip()
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        return f"{round(float(s), 2):.2f}"
    except Exception:
        return ""


# -----------------------------------------------------------------
# 🛒 EXPANDED CORPORATE PRODUCT LEXICON MATRIX
# -----------------------------------------------------------------
PRODUCT_LEXICON = {
    "Allsups 24 Case": [
        "ALLSUPS", "ALLSUPS 24", "ALLSUPS 24 CASE", "ALLSUP 24"
    ],
    "Food Club 24 Case": [
        "FOOD CLUB", "FOOD CLUB 24", "FOOD CLUB 24 CASE", "FC 24", "FOODCLUB 24", "FOODCLUB24 CASE", "FC24 CASE"
    ],
    "Food King 24 Case": [
        "FOOD KING 24", "FOOD KING 24 CASE", "FK 24", "FOODKING 24", "FOODKING24 CASE", "FK24 CASE"
    ],
    "Food King 40 Case": [
        "FOOD KING 40", "FOOD KING 40 CASE", "FK 40", "FOODKING 40", "FOODKING40 CASE", "FK40 CASE"
    ],
    "Juniors 24 Case": [
        "JUNIORS", "JUNIORS 24", "JUNIORS 24 CASE"
    ],
    "Lowes 24 Case": [
        "LOWES", "LOWES 24", "LOWES 24 CASE", "LOWES24"
    ],
    "Lowes 40 Case": [
        "LOWES 40", "LOWES 40 CASE"
    ],
    "Panhandle Pure 24 Case": [
        "PPL24", "PP 24 CASE", "PPL 24 CASSE", "PP24",
        "PANHANDLE PURE 24 CASE", "PPL 24 CASE", "PPL 24"
    ],
    "Panhandle Pure 40 Case": [
        "PPL40", "PP 40 CASE", "PPL 40 CASSE", "PP40",
        "PANHANDLE PURE 40 CASE", "PPL 40 CASE", "PPL 40"
    ],
    "Plains 24 Case": [
        "PLAINS", "PLAINS 24", "PLAINS 24 CASE"
    ],
    "Spring House 24 Case": [
        "SPRING HOUSE", "SPRING HOUSE 24", "SPRING HOUSE 24 CASE", "SH 24"
    ],
    "Toot N Totum 24 Case": [
        "TNT24", "TOOT N TOTUM", "TOOT 'N TOTUM 24", "TNT 24", 
        "TOOT N TOTUM 24 CASE", "TOOTN TOTUM 24 CASE"
    ],
}


def get_fuzzy_lexicon_match(value):
    if pd.isna(value):
        return None

    cleaned = clean_alphanumeric(value)
    if not cleaned:
        return None

    lookup = {}
    for standard_name, variants in PRODUCT_LEXICON.items():
        for variant in variants:
            lookup[clean_alphanumeric(variant)] = standard_name

    if cleaned in lookup:
        return lookup[cleaned]

    close = get_close_matches(cleaned, lookup.keys(), n=1, cutoff=0.82)
    return lookup[close[0]] if close else None


def _validate_columns(df, mapping, dataset_name):
    missing = [col for col in mapping.values() if col and col not in df.columns]
    if missing:
        raise KeyError(
            f"{dataset_name} is missing required column(s): {missing}. "
            f"Available columns: {list(df.columns)}"
        )


def execute_sales_reconciliation(df_qb, df_as400, qb_cols, as400_cols):
    df_qb = df_qb.copy()
    df_as400 = df_as400.copy()

    df_qb.columns = df_qb.columns.astype(str).str.strip().str.upper()
    df_as400.columns = df_as400.columns.astype(str).str.strip().str.upper()

    qb_required = {k: v for k, v in qb_cols.items() if k in ["po", "invoice", "amount"]}
    as400_required = {k: v for k, v in as400_cols.items() if k in ["po", "invoice", "amount"]}

    _validate_columns(df_qb, qb_required, "QuickBooks")
    _validate_columns(df_as400, as400_required, "AS400/Infinium")

    df_qb["_CLEAN_PO"] = df_qb[qb_cols["po"]].apply(clean_po)
    df_qb["_CLEAN_INV"] = df_qb[qb_cols["invoice"]].apply(clean_alphanumeric)
    df_qb["_NORM_AMT"] = df_qb[qb_cols["amount"]].apply(normalize_amount)

    df_as400["_CLEAN_PO"] = df_as400[as400_cols["po"]].apply(clean_po)
    df_as400["_CLEAN_INV"] = df_as400[as400_cols["invoice"]].apply(clean_alphanumeric)
    df_as400["_NORM_AMT"] = df_as400[as400_cols["amount"]].apply(normalize_amount)

    df_qb["_PO_KEY"] = df_qb["_CLEAN_PO"] + "|" + df_qb["_NORM_AMT"]
    df_qb["_INV_KEY"] = df_qb["_CLEAN_INV"] + "|" + df_qb["_NORM_AMT"]

    df_as400["_PO_KEY"] = df_as400["_CLEAN_PO"] + "|" + df_as400["_NORM_AMT"]
    df_as400["_INV_KEY"] = df_as400["_CLEAN_INV"] + "|" + df_as400["_NORM_AMT"]

    df_qb["IS_MATCHED"] = False
    df_qb["MATCH_METHOD"] = ""
    df_qb["MATCH_STATUS"] = "UNMATCHED QB"

    df_as400["IS_MATCHED"] = False
    df_as400["MATCH_STATUS"] = "UNMATCHED AS400"

    # Pass 1: PO + amount
    for idx, row in df_qb[(df_qb["_CLEAN_PO"] != "") & (df_qb["_NORM_AMT"] != "")].iterrows():
        if df_qb.at[idx, "IS_MATCHED"]:
            continue

        match_pool = df_as400[
            (df_as400["_PO_KEY"] == row["_PO_KEY"]) &
            (~df_as400["IS_MATCHED"])
        ]

        if not match_pool.empty:
            as400_idx = match_pool.index[0]
            df_qb.at[idx, "IS_MATCHED"] = True
            df_qb.at[idx, "MATCH_METHOD"] = "PO# + Amount"
            df_qb.at[idx, "MATCH_STATUS"] = "Matched"

            df_as400.at[as400_idx, "IS_MATCHED"] = True
            df_as400.at[as400_idx, "MATCH_STATUS"] = "Matched"

    # Pass 2: invoice + amount
    for idx, row in df_qb[
        (~df_qb["IS_MATCHED"]) &
        (df_qb["_CLEAN_INV"] != "") &
        (df_qb["_NORM_AMT"] != "")
    ].iterrows():
        match_pool = df_as400[
            (df_as400["_INV_KEY"] == row["_INV_KEY"]) &
            (~df_as400["IS_MATCHED"])
        ]

        if not match_pool.empty:
            as400_idx = match_pool.index[0]
            df_qb.at[idx, "IS_MATCHED"] = True
            df_qb.at[idx, "MATCH_METHOD"] = "Invoice # + Amount"
            df_qb.at[idx, "MATCH_STATUS"] = "Matched"

            df_as400.at[as400_idx, "IS_MATCHED"] = True
            df_as400.at[as400_idx, "MATCH_STATUS"] = "Matched"

    return df_qb, df_as400