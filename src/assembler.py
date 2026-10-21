"""
Assembler Module (PHA Server)
================================
Consolidates the 6 intermediate enriched files and Shared/ base data
into the Final Topic Report (46 columns).

Inputs:
  - input/Shared/PHA_Headings.xlsx (Standard target column order)
  - input/Shared/PHA_PQAH.xlsx (Base master population)
  - input/Shared/PHA_Direct_Spv.xlsx (Supervisor NIK & Name)
  - input/Shared/PHA_PA0105.xlsx (Email addresses)
  - input/Shared/PHA_Job_Layer.xlsx (Job Layer)
  - input/Shared/PHA_Employee_Level.xlsx (Employee Level mapping)
  - input/Shared/PHA_PA0021.xlsx (Number of Family Members)
  - Enriched intermediate files in output/intermediate/

Produces:
  - output/FINAL_HR_Topic_Report_YYYYMMDD.xlsx
  - logs/YYYYMMDD_HHMMSS_insight.xlsx (Audit integrity report)
"""

import logging
from datetime import datetime
from pathlib import Path
from typing import List, Optional
import pandas as pd

from src.common import normalize_id, strip_leading_zeros, load_excel, find_column_ci, normalize_date_column
from src.validator import validate_file_exists, validate_row_count
from src.excel_export import export_df
from src.insight_validator import validate_and_export_integrity_insight

SHARED_DIR       = Path("input/Shared")
HEADINGS_FILE    = SHARED_DIR / "PHA_Headings.xlsx"
PQAH_FILE        = SHARED_DIR / "PHA_PQAH.xlsx"

INTER_DIR        = Path("output/intermediate")
OUT_COST_CENTER  = INTER_DIR / "PHA_PA0027_Enriched_CostCenterDesc.xlsx"
OUT_JOIN_DATE    = INTER_DIR / "PQAH_Enriched_JoinDate_YoS.xlsx"
OUT_POS_TENURE   = INTER_DIR / "PQAH_Enriched_PositionTenure.xlsx"
OUT_CONTRACT     = INTER_DIR / "PHA_PA0016_With_ContractType_Description.xlsx"
OUT_EDUCATION    = INTER_DIR / "PHA_PA0022_HighestOrder_IgnoreTraining_WithBranchText.xlsx"
OUT_ENTITY       = INTER_DIR / "PHA_PQAH_With_Entity_DivisionDesc_Area_Function.xlsx"

OUTPUT_FINAL     = Path(f"output/PHA_FINAL_HR_Topic_Report_{datetime.now():%Y%m%d}.xlsx")


def _find_shared_file(candidates: List[str]) -> Optional[Path]:
    for name in candidates:
        p = SHARED_DIR / name
        if p.exists() and not p.name.startswith("~$") and p.is_file() and p.stat().st_size > 0:
            return p
    return None


def run(logger: logging.Logger) -> dict:
    module = "Final Topic Report Assembler"
    result = {
        "Module": module, "Input Rows": 0, "Output Rows": 0,
        "Matched": 0, "Unmatched": 0,
        "Status": "FAILED", "Output File": OUTPUT_FINAL.name,
    }

    try:
        headings_path = HEADINGS_FILE if HEADINGS_FILE.exists() else _find_shared_file(["PHA_Headings.xlsx", "IKP_Headings.xlsx"])
        pqah_path     = PQAH_FILE if PQAH_FILE.exists() else _find_shared_file(["PHA_PQAH.xlsx", "IKP_PQAH.xlsx"])

        validate_file_exists(headings_path, logger)
        validate_file_exists(pqah_path,     logger)

        headings_df = load_excel(headings_path)
        target_columns = list(headings_df.columns)
        logger.info(f"[Assembler] Target output template has {len(target_columns)} columns: {target_columns}")

        pqah = load_excel(pqah_path)
        input_rows = len(pqah)
        result["Input Rows"] = input_rows
        logger.info(f"[Assembler] Base PQAH has {input_rows:,} rows.")

        # In PQAH: "Personnel No." is NIK (join key), "Personnel Number" is Employee Name.
        pqah_pno_col = find_column_ci(pqah, "Personnel No.") or find_column_ci(pqah, "Personnel number") or pqah.columns[0]
        pqah["_pno_norm"] = normalize_id(pqah[pqah_pno_col])

        # Standardise common PQAH columns
        pqah_standard_maps = {
            "Personnel No.": ["Personnel No.", "Personnel number", "NIK"],
            "Personnel Number": ["Personnel Number", "Name", "Employee Name"],
            "Lvl": ["Lvl", "Level"],
            "CoCd": ["CoCd", "Company Code"],
            "Company Name": ["Company Name"],
            "ESgrp": ["ESgrp", "Employee Subgroup"],
            "Name of EE Subgroup": ["Name of EE Subgroup", "EE Subgroup Name"],
            "PArea": ["PArea", "Personnel Area"],
            "Personnel Area Text": ["Personnel Area Text"],
            "Subarea": ["Subarea", "Personnel Subarea"],
            "P. SubArea Text": ["P. SubArea Text", "Personnel Subarea Text"],
            "Position": ["Position", "Position ID"],
            "Position Name": ["Position Name", "Position Text"],
            "Org. Unit": ["Org. Unit", "Organizational Unit"],
            "Name of Organizational Unit": ["Name of Organizational Unit"],
            "Job": ["Job", "Job ID"],
            "Job Title": ["Job Title", "Job Text"],
            "Gender": ["Gender", "Gender text"],
            "Birth Date": ["Birth Date", "Date of Birth"],
            "Age": ["Age", "Age of employee"],
            "Nationality": ["Nationality"],
            "Birthplace": ["Birthplace"],
            "Religion": ["Religion"],
            "Marital Status": ["Marital Status", "MarSt"],
        }
        for tgt_col, src_candidates in pqah_standard_maps.items():
            for src in src_candidates:
                matched = find_column_ci(pqah, src)
                if matched:
                    pqah[tgt_col] = pqah[matched]
                    break

        # 1. Join Join Date & Year of Service
        if OUT_JOIN_DATE.exists():
            jd_df = load_excel(OUT_JOIN_DATE)
            jd_pno = find_column_ci(jd_df, "_pno_norm") or find_column_ci(jd_df, "Personnel No.") or find_column_ci(jd_df, "Personnel number")
            if jd_pno:
                jd_df["_pno_norm"] = normalize_id(jd_df[jd_pno])
                jd_cols = [c for c in ["Join Date", "Year of Service"] if find_column_ci(jd_df, c)]
                sub_jd = jd_df[["_pno_norm"] + jd_cols].drop_duplicates(subset=["_pno_norm"])
                pqah = pqah.merge(sub_jd, on="_pno_norm", how="left")
                logger.info(f"[Assembler] Merged Join Date / YoS columns: {jd_cols}")

        # 2. Join Position Effective Date & Current Tenure
        if OUT_POS_TENURE.exists():
            pt_df = load_excel(OUT_POS_TENURE)
            pt_pno = find_column_ci(pt_df, "_pno_norm") or find_column_ci(pt_df, "Personnel No.") or find_column_ci(pt_df, "Personnel number")
            if pt_pno:
                pt_df["_pno_norm"] = normalize_id(pt_df[pt_pno])
                tenure_col = find_column_ci(pt_df, "Current Tenure") or find_column_ci(pt_df, "Current Tenure (Years)")
                ped_col    = find_column_ci(pt_df, "Position Effective Date")
                sub_pt_dict = {"_pno_norm": pt_df["_pno_norm"]}
                if ped_col:
                    sub_pt_dict["Position Effective Date"] = pt_df[ped_col]
                if tenure_col:
                    sub_pt_dict["Current Tenure"] = pt_df[tenure_col]
                sub_pt = pd.DataFrame(sub_pt_dict).drop_duplicates(subset=["_pno_norm"])
                pqah = pqah.merge(sub_pt, on="_pno_norm", how="left")
                logger.info(f"[Assembler] Merged Position Tenure columns: {list(sub_pt_dict.keys())[1:]}")

        # 3. Join Entity, Division, Area, Function
        if OUT_ENTITY.exists():
            ent_df = load_excel(OUT_ENTITY)
            ent_pno = find_column_ci(ent_df, "_pno_norm") or find_column_ci(ent_df, "Personnel No.") or find_column_ci(ent_df, "Personnel number")
            if ent_pno:
                ent_df["_pno_norm"] = normalize_id(ent_df[ent_pno])
                sub_ent_dict = {"_pno_norm": ent_df["_pno_norm"]}
                for col in ["Entity", "Area", "Function"]:
                    matched = find_column_ci(ent_df, col)
                    if matched:
                        sub_ent_dict[col] = ent_df[matched]

                # Map "Grouping Division" from intermediate file to "Division" in final report
                grpdiv_col = find_column_ci(ent_df, "Grouping Division") or find_column_ci(ent_df, "Division")
                if grpdiv_col:
                    sub_ent_dict["Division"] = ent_df[grpdiv_col]

                sub_ent = pd.DataFrame(sub_ent_dict).drop_duplicates(subset=["_pno_norm"])
                pqah = pqah.merge(sub_ent, on="_pno_norm", how="left")

                # Special Rule 1: Override Entity for specific Personnel Numbers based on CoCd
                SPECIAL_ENTITY_PNOS = {"1062508", "01062508", "1137088", "01137088", "1142986", "01142986", "1142583", "01142583"}
                COCD_ENTITY_MAP = {"8188": "OKI", "8288": "PINDOKRW"}

                ent_col = find_column_ci(pqah, "Entity")
                cocd_col = find_column_ci(pqah, "CoCd")
                if ent_col and cocd_col:
                    pno_clean = strip_leading_zeros(pqah["_pno_norm"])
                    cocd_clean = normalize_id(pqah[cocd_col])
                    clean_special_pnos = {p.lstrip("0") for p in SPECIAL_ENTITY_PNOS}
                    mask = pno_clean.isin(clean_special_pnos) & cocd_clean.isin(COCD_ENTITY_MAP)
                    pqah.loc[mask, ent_col] = cocd_clean.map(COCD_ENTITY_MAP)
                    special_count = int(mask.sum())
                    if special_count > 0:
                        logger.info(f"[Assembler] Applied special Entity mapping for {special_count} employee(s) based on CoCd.")

                # Special Rule 2: Override Division to "GENERAL AFFAIR" for specific combination
                # Entity = "APP HEADQUARTERS", Division = "HR MILL, FORESTRY & OPERATIONS", Company Name = "Pindo Deli Pulp And Paper"
                div_col = find_column_ci(pqah, "Division")
                comp_col = find_column_ci(pqah, "Company Name")
                if ent_col and div_col and comp_col:
                    ent_s = pqah[ent_col].fillna("").astype(str).str.strip().str.upper()
                    div_s = pqah[div_col].fillna("").astype(str).str.strip().str.upper()
                    comp_s = pqah[comp_col].fillna("").astype(str).str.strip().str.upper()

                    mask_rule2 = (
                        (ent_s == "APP HEADQUARTERS") &
                        (div_s == "HUMAN RESOURCES") &
                        (comp_s == "PINDO DELI PULP AND PAPER")
                    )
                    pqah.loc[mask_rule2, div_col] = "GENERAL AFFAIR"
                    rule2_count = int(mask_rule2.sum())
                    if rule2_count > 0:
                        logger.info(f"[Assembler] Applied special Division override to 'GENERAL AFFAIR' for {rule2_count} employee(s).")

                merged_cols = [c for c in sub_ent_dict.keys() if c != "_pno_norm"]
                logger.info(f"[Assembler] Merged Entity/Division columns: {merged_cols}")

        # 4. Join Contract Type & Contract End Date (Sort Valid Until descending so most recent is first)
        out_contract_path = OUT_CONTRACT if OUT_CONTRACT.exists() else (INTER_DIR / "IKP_PA0016_With_ContractType_Description.xlsx")
        if out_contract_path.exists():
            ct_df = load_excel(out_contract_path)
            ct_pno = find_column_ci(ct_df, "Personnel number") or find_column_ci(ct_df, "Personnel No.") or find_column_ci(ct_df, "_pno_norm")
            if ct_pno:
                ct_df["_pno_norm"] = normalize_id(ct_df[ct_pno])
                ct_type_col = find_column_ci(ct_df, "Contract Type Description") or find_column_ci(ct_df, "Contract Type Text") or find_column_ci(ct_df, "Contract Type")
                ct_end_col  = find_column_ci(ct_df, "Valid Until") or find_column_ci(ct_df, "Contract End Date") or find_column_ci(ct_df, "End Date")

                # Sort Valid Until descending to prioritize the most recent contract
                if ct_end_col:
                    ct_df["_dt_sort"] = pd.to_datetime(ct_df[ct_end_col], dayfirst=True, errors="coerce")
                    ct_sorted = ct_df.sort_values(by="_dt_sort", ascending=False)
                else:
                    ct_sorted = ct_df

                ct_dedup = ct_sorted.drop_duplicates(subset=["_pno_norm"], keep="first")

                sub_ct_dict = {"_pno_norm": ct_dedup["_pno_norm"].values}
                if ct_type_col:
                    sub_ct_dict["Contract Type"] = ct_dedup[ct_type_col].values
                if ct_end_col:
                    sub_ct_dict["Contract End Date"] = ct_dedup[ct_end_col].values
                sub_ct = pd.DataFrame(sub_ct_dict)
                pqah = pqah.merge(sub_ct, on="_pno_norm", how="left")
                logger.info(f"[Assembler] Merged Contract columns (most recent Valid Until): {list(sub_ct_dict.keys())[1:]}")

        # 5. Join Education, Institute, Branch Study
        out_edu_path = OUT_EDUCATION if OUT_EDUCATION.exists() else (INTER_DIR / "IKP_IT0022_HighestOrder_IgnoreTraining_WithBranchText.xlsx")
        if out_edu_path.exists():
            edu_df = load_excel(out_edu_path)
            edu_pno = find_column_ci(edu_df, "Personnel number") or find_column_ci(edu_df, "Personnel No.") or find_column_ci(edu_df, "_pno_norm")
            if edu_pno:
                edu_df["_pno_norm"] = normalize_id(edu_df[edu_pno])
                edu_col  = find_column_ci(edu_df, "Education") or find_column_ci(edu_df, "Text")
                inst_col = find_column_ci(edu_df, "Institute") or find_column_ci(edu_df, "Institute/location") or find_column_ci(edu_df, "Name of establishment/institute")
                br_col   = find_column_ci(edu_df, "Branch Study") or find_column_ci(edu_df, "Branch of Study Text")
                sub_edu_dict = {"_pno_norm": edu_df["_pno_norm"]}
                if edu_col:
                    sub_edu_dict["Education"] = edu_df[edu_col]
                if inst_col:
                    sub_edu_dict["Institute"] = edu_df[inst_col]
                if br_col:
                    sub_edu_dict["Branch Study"] = edu_df[br_col]
                sub_edu = pd.DataFrame(sub_edu_dict).drop_duplicates(subset=["_pno_norm"])
                pqah = pqah.merge(sub_edu, on="_pno_norm", how="left")
                logger.info(f"[Assembler] Merged Education columns: {list(sub_edu_dict.keys())[1:]}")

        # 6. Join General Expense Cost Center, Salary Cost Center S4 & Descriptions
        out_cc_path = OUT_COST_CENTER
        if out_cc_path.exists():
            cc_df = load_excel(out_cc_path)
            cc_pno = find_column_ci(cc_df, "_pno_norm") or find_column_ci(cc_df, "Personnel Number") or find_column_ci(cc_df, "Personnel No.")
            if cc_pno:
                cc_df["_pno_norm"] = normalize_id(cc_df[cc_pno])
                cols_to_pull = [
                    "General Expense Cost Center", "General Expense Cost Center Desc",
                    "Salary Cost Center (S4)", "Salary Cost Center Desc (S4)"
                ]
                cols_map = {}
                for col_name in cols_to_pull:
                    found_col = find_column_ci(cc_df, col_name)
                    if found_col:
                        cols_map[found_col] = col_name

                if "General Expense Cost Center" not in cols_map.values():
                    old_cc = find_column_ci(cc_df, "Cost Center")
                    if old_cc:
                        cols_map[old_cc] = "General Expense Cost Center"
                if "General Expense Cost Center Desc" not in cols_map.values():
                    old_desc = find_column_ci(cc_df, "Cost Center Description")
                    if old_desc:
                        cols_map[old_desc] = "General Expense Cost Center Desc"

                if cols_map:
                    sub_cc = cc_df[["_pno_norm"] + list(cols_map.keys())].rename(columns=cols_map).drop_duplicates(subset=["_pno_norm"])
                    pqah = pqah.merge(sub_cc, on="_pno_norm", how="left")
                    logger.info(f"[Assembler] Merged Cost Center columns: {list(cols_map.values())}")

        # 7. Join Direct Supervisor (NIK & Name)
        spv_raw_df = None
        spv_path = _find_shared_file(["PHA_Direct_Spv.xlsx", "PHA_Direct Spv.xlsx", "IKP_Direct_Spv.xlsx", "IKP_Direct Spv.xlsx"])
        if spv_path:
            try:
                spv_df = load_excel(spv_path)
                spv_raw_df = spv_df.copy()
                spv_pno = find_column_ci(spv_df, "Personnel") or find_column_ci(spv_df, "Personnel No.") or find_column_ci(spv_df, "Personnel number") or spv_df.columns[0]
                
                # Strip leading zeros on BOTH sides to prevent 7-digit vs 8-digit SAP key mismatches
                spv_df["_pno_clean"] = strip_leading_zeros(normalize_id(spv_df[spv_pno]))
                pqah_pno_clean = strip_leading_zeros(pqah["_pno_norm"])

                dsup_col = find_column_ci(spv_df, "D.Superior") or find_column_ci(spv_df, "Superior NIK") or find_column_ci(spv_df, "Direct Superior NIK")
                name_col = find_column_ci(spv_df, "Direct Superior") or find_column_ci(spv_df, "Direct Sup") or find_column_ci(spv_df, "Superior Name")

                spv_dedup = spv_df.dropna(subset=["_pno_clean"]).drop_duplicates(subset=["_pno_clean"])

                # Direct map avoids any '_x' / '_y' suffix column collision
                if dsup_col:
                    nik_series = normalize_id(spv_dedup[dsup_col])
                    nik_lookup = pd.Series(nik_series.values, index=spv_dedup["_pno_clean"]).to_dict()
                    pqah["Superior NIK"] = pqah_pno_clean.map(nik_lookup).fillna("")
                else:
                    pqah["Superior NIK"] = ""

                if name_col:
                    name_series = spv_dedup[name_col].fillna("").astype(str).str.strip()
                    name_lookup = pd.Series(name_series.values, index=spv_dedup["_pno_clean"]).to_dict()
                    pqah["Superior Name"] = pqah_pno_clean.map(name_lookup).fillna("")
                else:
                    pqah["Superior Name"] = ""

                matched_count = int((pqah["Superior NIK"] != "").sum())
                logger.info(f"[Assembler] Merged Supervisor columns from {spv_path.name}: {matched_count}/{len(pqah)} matched.")
            except Exception as e:
                logger.warning(f"[Assembler] Could not merge Direct Supervisor from {spv_path.name}: {e}")
        else:
            logger.warning("[Assembler] Direct Supervisor file not found in Shared/.")

        # 8. Join Email (PA0105 / IT0105)
        pa0105_path = _find_shared_file(["PHA_PA0105.xlsx", "PHA_IT0105.xlsx"])
        if pa0105_path:
            try:
                pa0105_df = load_excel(pa0105_path)
                pno_105 = find_column_ci(pa0105_df, "Personnel number") or find_column_ci(pa0105_df, "Personnel No.") or pa0105_df.columns[0]
                email_col = find_column_ci(pa0105_df, "Long ID/Number") or find_column_ci(pa0105_df, "Email")
                if pno_105 and email_col:
                    pa0105_df["_pno_norm"] = normalize_id(pa0105_df[pno_105])
                    sub_105 = (
                        pa0105_df.dropna(subset=[email_col])
                                 .drop_duplicates(subset=["_pno_norm"])
                                 .set_index("_pno_norm")[email_col]
                    )
                    pqah["Email"] = pqah["_pno_norm"].map(sub_105)
                    logger.info(f"[Assembler] Merged Email from {pa0105_path.name}.")
            except Exception as e:
                logger.warning(f"[Assembler] Could not merge Email from {pa0105_path.name}: {e}")
        else:
            logger.warning("[Assembler] PA0105 file not found in Shared/.")

        # 9. Join Job Layer
        job_layer_path = _find_shared_file(["PHA_Job_Layer.xlsx", "IKP_Job_Layer.xlsx"])
        if job_layer_path:
            try:
                jl_df = load_excel(job_layer_path)
                job_key = find_column_ci(jl_df, "Job")
                layer_val = find_column_ci(jl_df, "Layer")
                pqah_job = find_column_ci(pqah, "Job")
                if job_key and layer_val and pqah_job:
                    jl_df["_job_norm"] = normalize_id(jl_df[job_key])
                    jl_lookup = jl_df.drop_duplicates(subset=["_job_norm"]).set_index("_job_norm")[layer_val]
                    pqah["_job_norm"] = normalize_id(pqah[pqah_job])
                    pqah["Layer"] = pqah["_job_norm"].map(jl_lookup)
                    pqah = pqah.drop(columns=["_job_norm"], errors="ignore")
                    logger.info(f"[Assembler] Merged Job Layer from {job_layer_path.name}.")
            except Exception as e:
                logger.warning(f"[Assembler] Could not merge Job Layer from {job_layer_path.name}: {e}")
        else:
            logger.warning("[Assembler] Job Layer file not found in Shared/.")

        # 10. Count and Join Number of Family Members (PA0021 / IT0021)
        pa0021_path = _find_shared_file(["PHA_PA0021.xlsx", "PHA_IT0021.xlsx", "IKP_IT0021.xlsx", "IKP_IT0021.xls"])
        if pa0021_path:
            try:
                pa0021_df = load_excel(pa0021_path)
                pa0021_pno = find_column_ci(pa0021_df, "Personnel number") or find_column_ci(pa0021_df, "Personnel No.") or pa0021_df.columns[0]
                pa0021_df["_pno_norm"] = normalize_id(pa0021_df[pa0021_pno])
                family_counts = (
                    pa0021_df.dropna(subset=["_pno_norm"])
                             .groupby("_pno_norm")
                             .size()
                             .reset_index(name="Number of Family Members")
                )
                pqah = pqah.merge(family_counts, on="_pno_norm", how="left")
                pqah["Number of Family Members"] = pqah["Number of Family Members"].fillna(0).astype(int)
                logger.info(f"[Assembler] Merged 'Number of Family Members' from {pa0021_path.name}.")
            except Exception as e:
                pqah["Number of Family Members"] = 0
                logger.warning(f"[Assembler] Failed to load/merge family members from {pa0021_path.name}: {e} — defaulted to 0.")
        else:
            pqah["Number of Family Members"] = 0
            logger.warning("[Assembler] Family members file (PA0021) not found in Shared/ — defaulted to 0.")

        # Ensure 'Number of Family Members' is in target_columns between 'Religion' and 'Marital Status'
        if "Number of Family Members" not in target_columns:
            if "Religion" in target_columns:
                rel_idx = target_columns.index("Religion")
                target_columns.insert(rel_idx + 1, "Number of Family Members")
            elif "Marital Status" in target_columns:
                ms_idx = target_columns.index("Marital Status")
                target_columns.insert(ms_idx, "Number of Family Members")
            else:
                target_columns.append("Number of Family Members")

        # 11. Construct Final DataFrame mapped to Headings exactly
        final_df = pd.DataFrame(index=range(len(pqah)))

        for col in target_columns:
            matched_col = find_column_ci(pqah, col)
            if matched_col:
                final_df[col] = pqah[matched_col].values
            else:
                final_df[col] = ""

        # Convert date columns to real Excel Date objects formatted as DD/MM/YYYY
        DATE_COLUMNS = [
            "Join Date", "Birth Date", "Contract End Date",
            "Position Effective Date", "Contract Start",
        ]
        active_date_cols = []
        for dcol in DATE_COLUMNS:
            matched_dcol = find_column_ci(final_df, dcol)
            if matched_dcol:
                norm_series = normalize_date_column(final_df[matched_dcol])
                final_df[matched_dcol] = pd.to_datetime(norm_series, format="%d/%m/%Y", errors="coerce")
                active_date_cols.append(matched_dcol)
                logger.info(f"[Assembler] Date column '{matched_dcol}' converted to Excel Date format.")

        validate_row_count(input_rows, len(final_df), module, logger)

        # 12. Data Integrity Insight Validation (Generates Excel report in logs/)
        validate_and_export_integrity_insight(final_df, spv_raw_df, logger)

        export_df(final_df, OUTPUT_FINAL, logger=logger, date_columns=active_date_cols)

        result.update({
            "Output Rows": len(final_df),
            "Matched": len(final_df),
            "Unmatched": 0,
            "Status": "SUCCESS",
        })
        logger.info(f"[Assembler] Final Topic Report successfully generated: {OUTPUT_FINAL.name} ({len(final_df):,} rows)")

    except Exception as e:
        logger.error(f"[Assembler] FAILED: {e}")

    return result