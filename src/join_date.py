"""
Automation 2: Join Date & Years of Service
============================================================
Input:  input/Shared/PHA_PQAH.xlsx
        input/Join Date & Year of Service/PHA_PA0041.xlsx
Output: output/intermediate/PQAH_Enriched_JoinDate_YoS.xlsx

Logic:
  - Auto-detect all "Date type" / "Date for date type" column PAIRS by name (case-insensitive)
  - Melt to long format: (Personnel number, date_type, date_value)
  - Normalize date type codes -> zero-padded 2-char ("1" -> "01")
  - Filter where date_type == "01"
  - Per employee: pick EARLIEST date (MIN)
  - LEFT JOIN onto PQAH by "Personnel No." (case-insensitive, exact name only)
  - Join Date = DD-Mon-YYYY (e.g. 27-Sep-2020)
  - Years of Service = (today - Join Date).days / 365.25, rounded 2dp
"""
import re
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime

from src.common import (normalize_id, normalize_date_type, load_excel,
                        get_required_column_ci)
from src.validator import (validate_file_exists, validate_columns,
                            validate_row_count, validate_dates)
from src.excel_export import export_df

SHARED_DIR  = Path("input/Shared")
PQAH_FILE   = SHARED_DIR / "PHA_PQAH.xlsx"
PA0041_FILE = Path("input/Join Date & Year of Service/PHA_PA0041.xlsx")
OUTPUT_FILE = Path("output/intermediate/PQAH_Enriched_JoinDate_YoS.xlsx")

PQAH_KEY   = "Personnel No."
PA0041_KEY = "Personnel number"

DATE_TYPE_PATTERN = re.compile(r"^date\s*type(\.\d+)?$", re.IGNORECASE)
DATE_VAL_PATTERN  = re.compile(r"^date\s*for\s*date\s*type(\.\d+)?$", re.IGNORECASE)


def _melt_date_pairs(PA0041: pd.DataFrame, PA0041_KEY: str, logger: logging.Logger) -> pd.DataFrame:
    dt_cols = sorted([c for c in PA0041.columns if DATE_TYPE_PATTERN.match(str(c).strip())])
    dv_cols = sorted([c for c in PA0041.columns if DATE_VAL_PATTERN.match(str(c).strip())])
    if not dt_cols:
        raise ValueError("[JoinDate] No 'Date type' columns found in PA0041.")
    logger.info(f"[JoinDate] {len(dt_cols)} Date type pair(s) detected")
    frames = []
    for dt_col, dv_col in zip(dt_cols, dv_cols):
        s = PA0041[[PA0041_KEY, dt_col, dv_col]].copy()
        s.columns = ["_pno_raw", "date_type", "date_value"]
        frames.append(s)
    return pd.concat(frames, ignore_index=True)


def run(logger: logging.Logger) -> dict:
    module = "Join Date & Years of Service"
    result = {
        "Module": module, "Input Rows": 0, "Output Rows": 0,
        "Matched": 0, "Unmatched": 0,
        "Status": "FAILED", "Output File": OUTPUT_FILE.name,
    }

    try:
        validate_file_exists(PQAH_FILE,   logger)
        validate_file_exists(PA0041_FILE, logger)

        pqah   = load_excel(PQAH_FILE)
        PA0041 = load_excel(PA0041_FILE)
        input_rows = len(pqah)
        result["Input Rows"] = input_rows
        logger.info(f"[JoinDate] PQAH: {input_rows:,} rows | PA0041: {len(PA0041):,} rows")

        validate_columns(pqah,   [PQAH_KEY],   "PHA_PQAH.xlsx",   logger)
        validate_columns(PA0041, [PA0041_KEY], "PHA_PA0041.xlsx", logger)

        pqah_key_col   = get_required_column_ci(pqah, PQAH_KEY)
        PA0041_KEY_col = get_required_column_ci(PA0041, PA0041_KEY)

        long_df = _melt_date_pairs(PA0041, PA0041_KEY_col, logger)
        long_df["_pno_norm"] = normalize_id(long_df["_pno_raw"])
        long_df["date_type"] = normalize_date_type(long_df["date_type"])

        # Filter Date type "01" only
        filtered = long_df[long_df["date_type"] == "01"].copy()
        logger.info(f"[JoinDate] Rows with Date type '01': {len(filtered):,}")

        filtered["date_value"] = pd.to_datetime(filtered["date_value"], errors="coerce")
        validate_dates(filtered["date_value"], "Date for date type (type 01)", logger)
        filtered = filtered.dropna(subset=["date_value"])

        # Earliest date per employee
        join_dates = (
            filtered.groupby("_pno_norm")["date_value"].min().reset_index()
        )
        join_dates.columns = ["_pno_norm", "_join_raw"]
        logger.info(f"[JoinDate] Employees with Join Date: {len(join_dates):,}")

        # LEFT JOIN onto PQAH strictly using "Personnel No." column
        pqah["_key"] = normalize_id(pqah[pqah_key_col])
        merged = pqah.merge(join_dates, left_on="_key", right_on="_pno_norm", how="left")
        merged = merged.drop(columns=["_pno_norm"], errors="ignore")

        # Calculate using today's date (dynamic per run)
        today = pd.Timestamp(datetime.today().date())
        merged["Year of Service"] = (
            (today - merged["_join_raw"]).dt.days / 365.25
        ).round(2)
        merged["Join Date"] = merged["_join_raw"].dt.strftime("%d/%m/%Y")
        merged = merged.drop(columns=["_key", "_join_raw"], errors="ignore")

        matched   = int(merged["Join Date"].notna().sum())
        unmatched = int(merged["Join Date"].isna().sum())
        logger.info(f"[JoinDate] Matched={matched:,}, Unmatched={unmatched:,}")

        validate_row_count(input_rows, len(merged), module, logger)
        export_df(merged, OUTPUT_FILE, logger)

        result.update({
            "Output Rows": len(merged), "Matched": matched,
            "Unmatched": unmatched, "Status": "SUCCESS",
        })
        logger.info("[JoinDate] Completed successfully.")

    except Exception as e:
        logger.error(f"[JoinDate] FAILED: {e}")

    return result
