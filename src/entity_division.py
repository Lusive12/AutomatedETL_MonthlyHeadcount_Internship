"""
Automation: Entity, Division, Area, Function (PHA Server)
============================================================
Input:  input/Entity, Division, Area, Function/PHA_PA9000.xlsx
        input/Shared/PHA_PQAH.xlsx
Output: output/intermediate/PHA_PQAH_With_Entity_DivisionDesc_Area_Function.xlsx

Logic:
  - Load PHA_PA9000.xlsx containing: Personnel Number, Entity, Division, Area, Function
  - Normalize Personnel Number on both files
  - Deduplicate PHA_PA9000 by latest record (End Date / Start Date descending)
  - Map Entity, Grouping Division, Area, Function onto PHA_PQAH
  - Preserve all PQAH rows (LEFT JOIN)
"""
import logging
from pathlib import Path
import pandas as pd

from src.common import normalize_id, load_excel, get_required_column_ci
from src.validator import validate_file_exists, validate_columns, validate_row_count
from src.excel_export import export_df

INPUT_DIR   = Path("input/Entity, Division, Area, Function")
PA9000_FILE = INPUT_DIR / "PHA_PA9000.xlsx"
PQAH_FILE   = Path("input/Shared/PHA_PQAH.xlsx")
OUTPUT_FILE = Path("output/intermediate/PHA_PQAH_With_Entity_DivisionDesc_Area_Function.xlsx")


def run(logger: logging.Logger) -> dict:
    module = "Entity, Division, Area, Function"
    result = {
        "Module": module, "Input Rows": 0, "Output Rows": 0,
        "Matched": 0, "Unmatched": 0,
        "Status": "FAILED", "Output File": OUTPUT_FILE.name,
    }

    try:
        validate_file_exists(PA9000_FILE, logger)
        validate_file_exists(PQAH_FILE,   logger)

        pa9000_df = load_excel(PA9000_FILE)
        pqah_df   = load_excel(PQAH_FILE)
        input_rows = len(pqah_df)
        result["Input Rows"] = input_rows
        logger.info(f"[EntityDivision] PA9000: {len(pa9000_df):,} rows | PQAH: {input_rows:,} rows")

        validate_columns(pa9000_df, ["Personnel Number", "Entity", "Grouping Division", "Area", "Function"], "PHA_PA9000.xlsx", logger)
        validate_columns(pqah_df,   ["Personnel No."], "PHA_PQAH.xlsx", logger)

        p9000_pno_col = get_required_column_ci(pa9000_df, "Personnel Number")
        pqah_pno_col  = get_required_column_ci(pqah_df, "Personnel No.")

        pa9000_df["_pno_norm"] = normalize_id(pa9000_df[p9000_pno_col])
        pqah_df["_pno_norm"]   = normalize_id(pqah_df[pqah_pno_col])

        # If duplicate records exist, sort by End Date and Start Date descending
        date_sort_cols = [c for c in ["End Date", "Start Date"] if c in pa9000_df.columns]
        if date_sort_cols:
            pa9000_df = pa9000_df.sort_values(by=date_sort_cols, ascending=False)

        master_map = pa9000_df.drop_duplicates(subset=["_pno_norm"]).set_index("_pno_norm")

        entity_col = get_required_column_ci(pa9000_df, "Entity")
        div_col    = get_required_column_ci(pa9000_df, "Grouping Division")
        area_col   = get_required_column_ci(pa9000_df, "Area")
        func_col   = get_required_column_ci(pa9000_df, "Function")

        pqah_df["Entity"]   = pqah_df["_pno_norm"].map(master_map[entity_col])
        pqah_df["Grouping Division"] = pqah_df["_pno_norm"].map(master_map[div_col])
        pqah_df["Area"]     = pqah_df["_pno_norm"].map(master_map[area_col])
        pqah_df["Function"] = pqah_df["_pno_norm"].map(master_map[func_col])

        matched   = int(pqah_df["Entity"].notna().sum())
        unmatched = int(pqah_df["Entity"].isna().sum())

        pqah_df = pqah_df.drop(columns=["_pno_norm"], errors="ignore")

        validate_row_count(input_rows, len(pqah_df), module, logger)
        export_df(pqah_df, OUTPUT_FILE, logger)

        result.update({
            "Output Rows": len(pqah_df), "Matched": matched,
            "Unmatched": unmatched, "Status": "SUCCESS",
        })
        logger.info(f"[EntityDivision] Completed successfully. Matched={matched:,}, Unmatched={unmatched:,}")

    except Exception as e:
        logger.error(f"[EntityDivision] FAILED: {e}")

    return result
