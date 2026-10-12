import difflib
import io
import re
import sqlite3
import openpyxl
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. DATABASE & PERSISTENCE LAYER (SQLite)
# ==========================================
DB_FILE = "quotation_master.db"


def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS master_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            main_material TEXT NOT NULL,
            sub_material TEXT NOT NULL UNIQUE,
            purchase_price REAL DEFAULT 0.0,
            discount REAL DEFAULT 0.0
        )
    """)
    conn.commit()
    return conn


def load_master_data():
    conn = get_db_connection()
    df = pd.read_sql_query(
        "SELECT id, main_material, sub_material, purchase_price, discount FROM master_items ORDER BY main_material, sub_material",
        conn,
    )
    conn.close()
    return df


def save_master_data(df):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM master_items")
    for _, row in df.iterrows():
        main_mat = str(row.get("main_material", "")).strip()
        sub_mat = str(row.get("sub_material", "")).strip()
        if not sub_mat:
            continue
        try:
            p_price = float(row.get("purchase_price", 0.0) or 0.0)
        except (ValueError, TypeError):
            p_price = 0.0
        try:
            disc = float(row.get("discount", 0.0) or 0.0)
        except (ValueError, TypeError):
            disc = 0.0

        cursor.execute(
            """
            INSERT OR REPLACE INTO master_items (main_material, sub_material, purchase_price, discount)
            VALUES (?, ?, ?, ?)
        """,
            (main_mat, sub_mat, p_price, disc),
        )
    conn.commit()
    conn.close()


def clear_master_db():
    conn = get_db_connection()
    conn.execute("DELETE FROM master_items")
    conn.commit()
    conn.close()


# ==========================================
# 2. FILE PARSING (EXCEL / XLSM / PDF)
# ==========================================
KNOWN_MAIN_CATEGORIES = [
    "CABLE GLANDS - POWER WITH SUITABLE RUBBER HOOD",
    "CABLE GLANDS -POWER WITH SUITABLE RUBBER HOOD",
    "CABLE GLANDS -POWER WITH RUBBER HOOD",
    "CABLE GLANDS - POWER",
    "CABLE GLANDS - SIGNAL",
    "CABLE GLAND - SIGNAL",
    "CABLE GLANDS  SIGNAL",
    "CABLE GLANDS SIGNAL",
    "REDUCER- ADAPTOR POWER",
    "REDUCER-ADAPTOR POWER",
    "REDUCER",
    "BLINDS POWER & SIGNAL",
    "BLIND - SIGNAL",
    "BLINDSIGNAL",
    "POWER CABLE LUGS & EARTHING LUGS (TINNED COPPER)",
    "LUGS POWER + EARTHING (TINNED COPPER)",
    "SIGNAL CABLE LUGS",
    "LUGS SIGNALS",
    "LUGS SIGNAL",
    "SLEEVES",
    "CABLE LABELS",
    "CABLE LABLES",
    "FERRULES",
    "CABLE TIE",
    "MISCELLANEOUS",
    "MISCELENIOUS",
]


def clean_text(val):
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def parse_master_from_excel(file_bytes):
    extracted_data = []
    wb_dict = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    for sheet_name, df in wb_dict.items():
        if "metadata" in sheet_name.lower():
            continue

        current_main = "MISCELLANEOUS"
        for _, row in df.iterrows():
            row_vals = [clean_text(v) for v in row if clean_text(v)]
            if not row_vals:
                continue

            first_col = row_vals[0]
            upper_first = first_col.upper()

            is_main = any(cat in upper_first for cat in KNOWN_MAIN_CATEGORIES)
            if is_main and (
                len(row_vals) <= 2
                or "POWER" in upper_first
                or "GLAND" in upper_first
            ):
                current_main = first_col
                continue

            desc = ""
            if (
                len(row_vals) >= 2
                and (
                    row_vals[0].isdigit()
                    or (len(row_vals[0]) <= 4 and "." in row_vals[0])
                )
            ):
                desc = row_vals[1]
            else:
                desc = row_vals[0]

            if (
                desc.lower()
                in [
                    "description",
                    "item",
                    "sr.no",
                    "particulars",
                    "part number",
                    "qty",
                    "remark",
                    "date",
                ]
                or desc.isdigit()
                or desc.startswith("ALFA LAVAL")
                or desc.startswith("PROJECT :")
                or desc.startswith("BILL OF MATERIAL")
            ):
                continue

            if len(desc) > 3:
                extracted_data.append(
                    {
                        "main_material": current_main,
                        "sub_material": desc,
                        "purchase_price": 0.0,
                        "discount": 0.0,
                    }
                )

    return pd.DataFrame(extracted_data)


def parse_master_from_pdf(file_bytes):
    extracted_data = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        current_main = "MISCELLANEOUS"
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                line_clean =