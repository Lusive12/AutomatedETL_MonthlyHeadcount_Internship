"""
Automation 1: General Cost Center & Salary Cost Center Enrichment (PHA Server)
================================================================================
Input:  input/General Cost Center & Cost Center Text/PHA_PA0027_01.xlsx  (General Expense Cost Center)
        input/General Cost Center & Cost Center Text/PHA_PA0027_02.xlsx  (Salary Cost Center S4)
        input/General Cost Center & Cost Center Text/PHA_CSKT_01.xlsx    (Cost Center Descriptions for PA0027_01)
        input/General Cost Center & Cost Center Text/PHA_CSKT_02.xlsx    (Cost Center Descriptions for PA0027_02)
        [Fallback: PHA_CSKT.xlsx if 01/02 not present]
Output: output/intermediate/PHA_PA0027_Enriched_CostCenterDesc.xlsx

Logic:
  - Build CSKT_01 lookup (for General Expense Cost Center): per Cost Center, keep row with MAX(Valid To) -> Description
  - Build CSKT_02 lookup (for Salary Cost Center S4): per Cost Center, keep row with MAX(Valid To) -> Description
  - PHA_PA0027_01:
      Cost Center -> "General Expense Cost Center"
      Lookup (CSKT_01) -> "General Expense Cost Center Desc"
  - PHA_PA0027_02:
      Cost Center -> "Salary Cost Center (S4)"
      Lookup (CSKT_02) -> "Salary Cost Center Desc (S4)"
  - Merge both on normalized Personnel Number and export to intermediate
"""
import re
import logging
import pandas as pd
from pathlib import Path

from src.common import normalize_id, load_excel, get_required_column_ci
from src.validator import (validate_file_exists, validate_columns,
                            validate_row_count)
from src.excel_export import export_df

INPUT_DIR        = Path("input/General Cost Center & Cost Center Text")
PA0027_01_FILE   = INPUT_DIR / "PHA_PA0027_01.xlsx"
PA0027_02_FILE   = INPUT_DIR / "PHA_PA0027_02.xlsx"
CSKT_01_FILE     = INPUT_DIR / "PHA_CSKT_01.xlsx"
CSKT_02_FILE     = INPUT_DIR / "PHA_CSKT_02.xlsx"
CSKT_LEGACY_FILE = INPUT_DIR / "PHA_CSKT.xlsx"
OUTPUT_FILE      = Path("output/intermediate/PHA_PA0027_Enriched_CostCenterDesc.xlsx")

CSKT_REQUIRED    = ["Cost Center", "Valid To", "Description"]


def _build_cskt_lookup(cskt_df: pd.DataFrame, file_name: str, logger: logging.Logger) -> pd.Series:
    """Extract unique Cost Center -> Description mapping based on MAX(Valid To)."""
    validate_columns(cskt_df, CSKT_REQUIRED, file_name, logger)
    cskt_cc_col    = get_required_column_ci(cskt_df, "Cost Center")
    cskt_valto_col = get_required_column_ci(cskt_df, "Valid To")
    cskt_desc_col  = get_required_column_ci(cskt_df, "Description")

    df = cskt_df.copy()
    df["_cc_norm"] = normalize_id(df[cskt_cc_col])
    df["_val_to"]  = pd.to_datetime(df[cskt_valto_col], errors="coerce")
    lookup = (
        df.sort_values("_val_to", ascending=False)
          .drop_duplicates(subset=["_cc_norm"], keep="first")
          .set_index("_cc_norm")[cskt_desc_col]
    )
    logger.info(f"[CostCenter] Lookup built from {file_name}: {len(lookup):,} unique CCs")
    return lookup


def run(logger: logging.Logger) -> dict:
    module = "General Cost Center"
    result = {
        "Module": module, "Input Rows": 0, "Output Rows": 0,
        "Matched": 0, "Unmatched": 0,
        "Status": "FAILED", "Output File": OUTPUT_FILE.name,
    }

    try:
        validate_file_exists(PA0027_01_FILE, logger)
        validate_file_exists(PA0027_02_FILE, logger)

        # Resolve CSKT files with fallback to legacy PHA_CSKT.xlsx
        cskt_01_path = CSKT_01_FILE if CSKT_01_FILE.exists() else CSKT_LEGACY_FILE
        cskt_02_path = CSKT_02_FILE if CSKT_02_FILE.exists() else CSKT_LEGACY_FILE

        validate_file_exists(cskt_01_path, logger)
        validate_file_exists(cskt_02_path, logger)

        pa0027_01 = load_excel(PA0027_01_FILE)
        pa0027_02 = load_excel(PA0027_02_FILE)
        cskt_01   = load_excel(cskt_01_path)
        cskt_02   = load_excel(cskt_02_path) if cskt_02_path != cskt_01_path else cskt_01

        input_rows = len(pa0027_01)
        result["Input Rows"] = input_rows
        logger.info(
            f"[CostCenter] PA0027_01: {len(pa0027_01):,} rows | PA0027_02: {len(pa0027_02):,} rows | "
            f"CSKT_01 ({cskt_01_path.name}): {len(cskt_01):,} rows | CSKT_02 ({cskt_02_path.name}): {len(cskt_02):,} rows"
        )

        validate_columns(pa0027_01, ["Personnel Number", "Cost Center"], "PHA_PA0027_01.xlsx", logger)
        validate_columns(pa0027_02, ["Personnel Number", "Cost Center"], "PHA_PA0027_02.xlsx", logger)

        # Build lookups separately
        cskt_01_lookup = _build_cskt_lookup(cskt_01, cskt_01_path.name, logger)
        cskt_02_lookup = _build_cskt_lookup(cskt_02, cskt_02_path.name, logger)

        # Process PA0027_01 -> General Expense Cost Center (linked with CSKT_01)
        pno_col_01 = get_required_column_ci(pa0027_01, "Personnel Number")
        cc_col_01  = get_required_column_ci(pa0027_01, "Cost Center")

        pno_norm_01 = normalize_id(pa0027_01[pno_col_01])
        cc_norm_01  = normalize_id(pa0027_01[cc_col_01])
        sub1 = pd.DataFrame({
            "_pno_norm": pno_norm_01,
            "Personnel Number": pa0027_01[pno_col_01],
            "General Expense Cost Center": pa0027_01[cc_col_01],
            "General Expense Cost Center Desc": cc_norm_01.map(cskt_01_lookup),
        }).drop_duplicates(subset=["_pno_norm"])

        # Process PA0027_02 -> Salary Cost Center (S4) (linked with CSKT_02)
        pno_col_02 = get_required_column_ci(pa0027_02, "Personnel Number")
        cc_col_02  = get_required_column_ci(pa0027_02, "Cost Center")

        pno_norm_02 = normalize_id(pa0027_02[pno_col_02])
        cc_norm_02  = normalize_id(pa0027_02[cc_col_02])
        sub2 = pd.DataFrame({
            "_pno_norm": pno_norm_02,
            "Salary Cost Center (S4)": pa0027_02[cc_col_02],
            "Salary Cost Center Desc (S4)": cc_norm_02.map(cskt_02_lookup),
        }).drop_duplicates(subset=["_pno_norm"])

        # Merge both cost center frames
        result_df = pd.merge(sub1, sub2, on="_pno_norm", how="outer")
        if "Personnel Number" not in result_df.columns:
            result_df["Personnel Number"] = result_df["_pno_norm"]

        total_matched = int(result_df["General Expense Cost Center Desc"].notna().sum() +
                            result_df["Salary Cost Center Desc (S4)"].notna().sum())
        total_unmatched = int(result_df["General Expense Cost Center Desc"].isna().sum() +
                              result_df["Salary Cost Center Desc (S4)"].isna().sum())

        export_df(result_df, OUTPUT_FILE, logger)

        result.update({
            "Output Rows": len(result_df), "Matched": total_matched,
            "Unmatched": total_unmatched, "Status": "SUCCESS",
        })
        logger.info(f"[CostCenter] Completed successfully. Saved: {OUTPUT_FILE.name} ({len(result_df):,} rows)")

    except Exception as e:
        logger.error(f"[CostCenter] FAILED: {e}")

    return result