"""
Automation 3: Position Effective Date & Current Tenure (PHA Server)
=====================================================================
Input:  input/Shared/PHA_PQAH.xlsx
        input/Position Effective Date & Current Tenure/PHA_HRP1001.xlsx
Output: output/intermediate/PQAH_Enriched_PositionTenure.xlsx

Logic:
  - Build composite key: normalize(Personnel No.) + "|" + normalize(Position)
  - In HRP1001: strip leading zeros from "ID of related object"
  - Match "Start Date" and "Object ID"
  - Select MIN(Start Date) per composite key
  - LEFT JOIN onto PQAH
  - Current Tenure (Years) = (today - Earliest Start Date).days / 365.25, round 2dp
"""
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime

from src.common import (normalize_id, strip_leading_zeros, load_excel,
                        get_required_column_ci)
from src.validator import (validate_file_exists, validate_columns,
                            validate_row_count, validate_dates)
from src.excel_export import export_df

SHARED_DIR   = Path("input/Shared")
PQAH_FILE    = SHARED_DIR / "PHA_PQAH.xlsx"
HRP1001_FILE = Path("input/Position Effective Date & Current Tenure/PHA_HRP1001.xlsx")
OUTPUT_FILE  = Path("output/intermediate/PQAH_Enriched_PositionTenure.xlsx")

PQAH_REQUIRED    = ["Personnel No.", "Position"]
HRP1001_REQUIRED = ["Object ID", "ID of related object", "Start Date"]


def run(logger: logging.Logger) -> dict:
    module = "Position Effective Date & Current Tenure"
    result = {
        "Module": module, "Input Rows": 0, "Output Rows": 0,
        "Matched": 0, "Unmatched": 0,
        "Status": "FAILED", "Output File": OUTPUT_FILE.name,
    }

    try:
        validate_file_exists(PQAH_FILE,    logger)
        validate_file_exists(HRP1001_FILE, logger)

        pqah    = load_excel(PQAH_FILE)
        hrp1001 = load_excel(HRP1001_FILE)
        input_rows = len(pqah)
        result["Input Rows"] = input_rows
        logger.info(f"[PosTenure] PQAH: {input_rows:,} rows | HRP1001: {len(hrp1001):,} rows")

        validate_columns(pqah,    PQAH_REQUIRED,    "PHA_PQAH.xlsx",    logger)
        validate_columns(hrp1001, HRP1001_REQUIRED, "PHA_HRP1001.xlsx", logger)

        pno_col   = get_required_column_ci(pqah, "Personnel No.")
        pos_col   = get_required_column_ci(pqah, "Position")
        objid_col = get_required_column_ci(hrp1001, "Object ID")
        relid_col = get_required_column_ci(hrp1001, "ID of related object")
        sdate_col = get_required_column_ci(hrp1001, "Start Date")

        hrp1001["_sdate"] = pd.to_datetime(hrp1001[sdate_col], errors="coerce")
        validate_dates(hrp1001["_sdate"], "Start Date", logger)
        hrp1001 = hrp1001.dropna(subset=["_sdate"]).copy()

        # Composite key: obj_id + "|" + stripped(rel_id)
        hrp1001["_comp_key"] = (
            normalize_id(hrp1001[objid_col])
            + "|"
            + strip_leading_zeros(hrp1001[relid_col])
        )

        min_dates = (
            hrp1001.groupby("_comp_key")["_sdate"].min().reset_index()
        )
        min_dates.columns = ["_comp_key", "_pos_start_raw"]
        logger.info(f"[PosTenure] Unique composite keys in HRP1001: {len(min_dates):,}")

        # PQAH composite key: Position + "|" + stripped(Personnel No.)
        pqah["_comp_key"] = (
            normalize_id(pqah[pos_col])
            + "|"
            + strip_leading_zeros(pqah[pno_col])
        )

        merged = pqah.merge(min_dates, on="_comp_key", how="left")
        merged = merged.drop(columns=["_comp_key"], errors="ignore")

        today = pd.Timestamp(datetime.today().date())
        merged["Current Tenure (Years)"] = (
            (today - merged["_pos_start_raw"]).dt.days / 365.25
        ).round(2)
        merged["Position Effective Date"] = (
            merged["_pos_start_raw"].dt.strftime("%d/%m/%Y")
        )
        merged = merged.drop(columns=["_pos_start_raw"], errors="ignore")

        matched   = int(merged["Position Effective Date"].notna().sum())
        unmatched = int(merged["Position Effective Date"].isna().sum())
        logger.info(f"[PosTenure] Matched={matched:,}, Unmatched={unmatched:,}")

        validate_row_count(input_rows, len(merged), module, logger)
        export_df(merged, OUTPUT_FILE, logger)

        result.update({
            "Output Rows": len(merged), "Matched": matched,
            "Unmatched": unmatched, "Status": "SUCCESS",
        })
        logger.info("[PosTenure] Completed successfully.")

    except Exception as e:
        logger.error(f"[PosTenure] FAILED: {e}")

    return result
