"""
Data Integrity & Validation Engine (PHA Server)
============================================================
Checks data integrity rules and generates an audit insight report
saved in the logs/ folder.
"""
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional
import pandas as pd
from src.common import find_column_ci
from src.excel_export import export_df


def validate_and_export_integrity_insight(
    final_df: pd.DataFrame,
    spv_df: Optional[pd.DataFrame],
    logger: Optional[logging.Logger] = None,
    log_folder: str = "logs"
) -> Path:
    issues = []

    # 1. Validation on Direct Supervisor Source Data
    if spv_df is not None and not spv_df.empty:
        pno_col  = find_column_ci(spv_df, "Personnel") or find_column_ci(spv_df, "Personnel No.")
        dsup_col = find_column_ci(spv_df, "D.Superior") or find_column_ci(spv_df, "Superior NIK")
        name_col = find_column_ci(spv_df, "Direct Sup") or find_column_ci(spv_df, "Superior Name")

        if pno_col and (dsup_col or name_col):
            for _, row in spv_df.iterrows():
                pno = row[pno_col]
                dsup_val = str(row[dsup_col]).strip() if (dsup_col and pd.notna(row[dsup_col])) else ""
                name_val = str(row[name_col]).strip() if (name_col and pd.notna(row[name_col])) else ""

                dsup_empty = (dsup_val == "" or dsup_val.lower() in ("nan", "none"))
                name_empty = (name_val == "" or name_val.lower() in ("nan", "none"))

                if dsup_empty and name_empty:
                    issues.append({
                        "Personnel No.": pno,
                        "Sumber Data": "PHA_Direct_Spv.xlsx",
                        "Kategori Masalah": "Atasan Langsung Kosong",
                        "Deskripsi Masalah": "Kolom 'D.Superior' (NIK Atasan) dan 'Direct Sup' (Nama Atasan) keduanya kosong.",
                    })
                elif dsup_empty:
                    issues.append({
                        "Personnel No.": pno,
                        "Sumber Data": "PHA_Direct_Spv.xlsx",
                        "Kategori Masalah": "Atasan Langsung Kosong",
                        "Deskripsi Masalah": "Kolom 'D.Superior' (NIK Atasan) kosong.",
                    })
                elif name_empty:
                    issues.append({
                        "Personnel No.": pno,
                        "Sumber Data": "PHA_Direct_Spv.xlsx",
                        "Kategori Masalah": "Atasan Langsung Kosong",
                        "Deskripsi Masalah": "Kolom 'Direct Sup' (Nama Atasan) kosong.",
                    })

    # 2. Validation on Final Report
    pno_col           = find_column_ci(final_df, "Personnel No.") or find_column_ci(final_df, "Personnel Number")
    subarea_col       = find_column_ci(final_df, "Subarea")
    cost_center_col   = find_column_ci(final_df, "General Expense Cost Center")
    layer_col         = find_column_ci(final_df, "Layer")
    subgroup_col      = find_column_ci(final_df, "Name of EE Subgroup")
    contract_type_col = find_column_ci(final_df, "Contract Type")
    contract_end_col  = find_column_ci(final_df, "Contract End Date")

    for idx, row in final_df.iterrows():
        pno = row[pno_col] if pno_col else f"Baris_{idx+1}"

        # Rule 1: 4 digit Subarea == 4 digit terakhir General Expense Cost Center
        if subarea_col and cost_center_col:
            sub_val = str(row[subarea_col]).strip() if pd.notna(row[subarea_col]) else ""
            cc_raw = str(row[cost_center_col]).strip() if pd.notna(row[cost_center_col]) else ""
            if cc_raw.endswith(".0"):
                cc_raw = cc_raw[:-2]

            if not cc_raw or cc_raw.lower() in ("nan", "none"):
                issues.append({
                    "Personnel No.": pno,
                    "Sumber Data": "Final Report",
                    "Kategori Masalah": "Subarea vs Cost Center Mismatch",
                    "Deskripsi Masalah": f"Kolom 'General Expense Cost Center' kosong, tidak dapat dicocokkan dengan Subarea '{sub_val}'.",
                })
            else:
                cc_last4 = cc_raw[-4:] if len(cc_raw) >= 4 else cc_raw
                if sub_val.upper() != cc_last4.upper():
                    issues.append({
                        "Personnel No.": pno,
                        "Sumber Data": "Final Report",
                        "Kategori Masalah": "Subarea vs Cost Center Mismatch",
                        "Deskripsi Masalah": f"4 angka Subarea ('{sub_val}') tidak sama dengan 4 digit terakhir General Expense Cost Center ('{cc_last4}' dari '{cc_raw}').",
                    })

        # Rule 2: Kolom Layer harus terisi
        if layer_col:
            layer_val = str(row[layer_col]).strip() if pd.notna(row[layer_col]) else ""
            if layer_val == "" or layer_val.lower() in ("nan", "none"):
                issues.append({
                    "Personnel No.": pno,
                    "Sumber Data": "Final Report",
                    "Kategori Masalah": "Layer Kosong",
                    "Deskripsi Masalah": "Kolom 'Layer' kosong / tidak terisi.",
                })

        # Rule 3: Subgroup 'Contract' / 'Kontrak' wajib punya Contract Type & Contract End Date
        if subgroup_col:
            sg_val = str(row[subgroup_col]).strip() if pd.notna(row[subgroup_col]) else ""
            if "CONTRACT" in sg_val.upper() or "KONTRAK" in sg_val.upper():
                ct_val = str(row[contract_type_col]).strip() if (contract_type_col and pd.notna(row[contract_type_col])) else ""
                ce_val = str(row[contract_end_col]).strip() if (contract_end_col and pd.notna(row[contract_end_col])) else ""

                ct_empty = (ct_val == "" or ct_val.lower() in ("nan", "none"))
                ce_empty = (ce_val == "" or ce_val.lower() in ("nan", "none"))

                if ct_empty and ce_empty:
                    issues.append({
                        "Personnel No.": pno,
                        "Sumber Data": "Final Report",
                        "Kategori Masalah": "Data Kontrak Tidak Lengkap",
                        "Deskripsi Masalah": f"Name of EE Subgroup adalah '{sg_val}', tetapi 'Contract Type' dan 'Contract End Date' keduanya kosong.",
                    })
                elif ct_empty:
                    issues.append({
                        "Personnel No.": pno,
                        "Sumber Data": "Final Report",
                        "Kategori Masalah": "Data Kontrak Tidak Lengkap",
                        "Deskripsi Masalah": f"Name of EE Subgroup adalah '{sg_val}', tetapi 'Contract Type' kosong.",
                    })
                elif ce_empty:
                    issues.append({
                        "Personnel No.": pno,
                        "Sumber Data": "Final Report",
                        "Kategori Masalah": "Data Kontrak Tidak Lengkap",
                        "Deskripsi Masalah": f"Name of EE Subgroup adalah '{sg_val}', tetapi 'Contract End Date' kosong.",
                    })

    # 3. Export to Excel in logs/ folder
    if issues:
        insight_df = pd.DataFrame(issues)
    else:
        insight_df = pd.DataFrame([{
            "Personnel No.": "-",
            "Sumber Data": "Semua Sumber Data",
            "Kategori Masalah": "Semua Validasi Lolos",
            "Deskripsi Masalah": "Integritas data 100% valid. Tidak ditemukan data kosong atau tidak cocok.",
        }])

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if logger and logger.handlers:
        for h in logger.handlers:
            if isinstance(h, logging.FileHandler) and h.baseFilename:
                timestamp = Path(h.baseFilename).stem
                break

    logs_dir = Path(log_folder)
    logs_dir.mkdir(parents=True, exist_ok=True)
    insight_path = logs_dir / f"{timestamp}_insight.xlsx"

    try:
        export_df(insight_df, insight_path, logger=None)
        if logger:
            logger.info(
                f"[InsightValidator] Data integrity check complete: {len(issues)} issue(s) flagged. "
                f"Saved to: {insight_path.name}"
            )
    except Exception as e:
        if logger:
            logger.error(f"[InsightValidator] Failed to export insight excel: {e}")

    return insight_path
