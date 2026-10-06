"""
Automation: Education, Institute & Branch of Study (PHA Server)
================================================================
Input:  input/Education, Institute, Branch of Study/PHA_PA0022.xlsx
        input/Education, Institute, Branch of Study/Order_PHA_PA0022.xlsx
        input/Education, Institute, Branch of Study/PHA_Branch of Study.xlsx
Output: output/intermediate/PHA_PA0022_HighestOrder_IgnoreTraining_WithBranchText.xlsx

Logic:
  Step 1: Join Order table onto PA0022.
          "Educational est." in PA0022 == "Key" in Order_PHA_PA0022.xlsx
          Educational est. values: "2", "3" etc. -> zero-pad to "02", "03" to match Key.
          This brings in: Key, Order, Text

  Step 2: Exclude training records:
          - Key == "11" (Training key)
          - Order == 11 (Training order as int)

  Step 3: Per employee (Personnel number), sort by Order DESCENDING (numeric),
          keep first row = highest qualification.

  Step 4: BranchCodeNorm = Branch of Study 1 zero-padded to 5 digits
          e.g. "40" -> "00040"

  Step 5: Lookup Branch of Study Text from PHA_Branch of Study.xlsx:
          BranchCodeNorm == Branch of Study column in lookup -> pull Branch of Study Text
"""
import logging
from pathlib import Path
import pandas as pd

from src.common import (normalize_id, load_excel, get_required_column_ci,
                        find_column_ci)
from src.validator import (validate_file_exists, validate_columns)
from src.excel_export import export_df

INPUT_DIR   = Path("input/Education, Institute, Branch of Study")
PA0022_FILE = INPUT_DIR / "PHA_PA0022.xlsx"
ORDER_FILE  = INPUT_DIR / "Order_PHA_PA0022.xlsx"
BRANCH_FILE = INPUT_DIR / "PHA_Branch of Study.xlsx"
OUTPUT_FILE = Path("output/intermediate/PHA_PA0022_HighestOrder_IgnoreTraining_WithBranchText.xlsx")

TRAINING_KEY   = "11"   # Key value that means Training
TRAINING_ORDER = 11     # Order value that means Training


def _zfill2(s: str) -> str:
    """Zero-pad a numeric string to 2 digits. '2' -> '02', '10' -> '10'."""
    s = s.strip()
    if s.isdigit():
        return s.zfill(2)
    return s


def run(logger: logging.Logger) -> dict:
    module = "Education, Institute & Branch of Study"
    result = {
        "Module": module, "Input Rows": 0, "Output Rows": 0,
        "Matched": 0, "Unmatched": 0,
        "Status": "FAILED", "Output File": OUTPUT_FILE.name,
    }

    try:
        validate_file_exists(PA0022_FILE, logger)
        validate_file_exists(ORDER_FILE,  logger)
        validate_file_exists(BRANCH_FILE, logger)

        pa0022 = load_excel(PA0022_FILE)
        order  = load_excel(ORDER_FILE)
        branch = load_excel(BRANCH_FILE)

        input_rows = len(pa0022)
        result["Input Rows"] = input_rows
        logger.info(f"[Education] PA0022: {input_rows:,} rows | Order: {len(order):,} rows | Branch: {len(branch):,} rows")

        # ── Validate required columns ─────────────────────────────────────────
        validate_columns(pa0022, ["Personnel number", "Educational est.", "Branch of Study 1"],
                         "PHA_PA0022.xlsx", logger)
        validate_columns(order,  ["Key", "Order", "Text"],
                         "Order_PHA_PA0022.xlsx", logger)
        validate_columns(branch, ["Branch of Study", "Branch of Study Text"],
                         "PHA_Branch of Study.xlsx", logger)

        pno_col     = get_required_column_ci(pa0022, "Personnel number")
        edu_est_col = get_required_column_ci(pa0022, "Educational est.")
        branch1_col = get_required_column_ci(pa0022, "Branch of Study 1")

        ord_key_col  = get_required_column_ci(order, "Key")
        ord_ord_col  = get_required_column_ci(order, "Order")
        ord_text_col = get_required_column_ci(order, "Text")

        br_key_col  = get_required_column_ci(branch, "Branch of Study")
        br_text_col = get_required_column_ci(branch, "Branch of Study Text")

        # ── Step 1: Normalize keys then join ──────────────────────────────────
        pa0022["_edu_zfill"] = (
            normalize_id(pa0022[edu_est_col])
            .apply(_zfill2)
        )

        order["_ord_key_norm"] = normalize_id(order[ord_key_col]).apply(_zfill2)
        order["_order_num"] = pd.to_numeric(
            normalize_id(order[ord_ord_col]), errors="coerce"
        ).fillna(0).astype(int)

        logger.info(f"[Education] Order keys: {sorted(order['_ord_key_norm'].tolist())}")

        order_dedup    = order.drop_duplicates(subset=["_ord_key_norm"]).set_index("_ord_key_norm")
        order_num_map  = order_dedup["_order_num"]
        order_text_map = order_dedup[ord_text_col]

        pa0022["Key"]   = pa0022["_edu_zfill"].where(pa0022["_edu_zfill"].isin(order_dedup.index))
        pa0022["Order"] = pa0022["_edu_zfill"].map(order_num_map)
        pa0022["Text"]  = pa0022["_edu_zfill"].map(order_text_map)

        matched_order = int(pa0022["Key"].notna().sum())
        logger.info(f"[Education] Order join: {matched_order:,}/{input_rows:,} rows matched")

        # ── Step 2: Exclude training records ──────────────────────────────────
        mask_training = (
            (pa0022["Key"] == TRAINING_KEY) |
            (pa0022["Order"] == TRAINING_ORDER)
        )
        excluded = int(mask_training.sum())
        pa0022_filtered = pa0022[~mask_training].copy()
        logger.info(
            f"[Education] Excluded {excluded:,} training records "
            f"(Key={TRAINING_KEY!r} or Order={TRAINING_ORDER}). "
            f"Remaining: {len(pa0022_filtered):,}"
        )

        # ── Step 3: Keep highest Order per employee ───────────────────────────
        pa0022_filtered["_pno_norm"] = normalize_id(pa0022_filtered[pno_col])
        pa0022_filtered = pa0022_filtered.sort_values(
            by=["_pno_norm", "Order"], ascending=[True, False]
        )
        highest_edu = pa0022_filtered.drop_duplicates(
            subset=["_pno_norm"], keep="first"
        ).copy()
        logger.info(f"[Education] Unique employees with highest education: {len(highest_edu):,}")

        # ── Step 4: BranchCodeNorm = zero-pad Branch of Study 1 to 5 digits ──
        highest_edu["_br1_norm"] = normalize_id(highest_edu[branch1_col])
        highest_edu["BranchCodeNorm"] = highest_edu["_br1_norm"].apply(
            lambda x: x.zfill(5) if x.isdigit() else ""
        )

        # ── Step 5: Lookup Branch of Study Text ───────────────────────────────
        branch["_br_key_norm"] = normalize_id(branch[br_key_col])
        br_lookup = (
            branch.drop_duplicates(subset=["_br_key_norm"])
                  .set_index("_br_key_norm")[br_text_col]
        )
        highest_edu["Branch of Study Text"] = highest_edu["BranchCodeNorm"].map(br_lookup)

        matched = int(highest_edu["Text"].notna().sum())
        unmatched = int(highest_edu["Text"].isna().sum())

        output_cols = [
            c for c in pa0022.columns
            if c not in ("_edu_zfill", "Key", "Order", "Text")
        ] + ["Key", "Order", "Text", "BranchCodeNorm", "Branch of Study Text"]
        output_cols = [c for c in output_cols if c in highest_edu.columns]
        result_df = highest_edu[output_cols]

        export_df(result_df, OUTPUT_FILE, logger)

        result.update({
            "Output Rows": len(result_df), "Matched": matched,
            "Unmatched": unmatched, "Status": "SUCCESS",
        })
        logger.info(f"[Education] Completed successfully. Saved: {OUTPUT_FILE.name}")

    except Exception as e:
        logger.error(f"[Education] FAILED: {e}")

    return result
