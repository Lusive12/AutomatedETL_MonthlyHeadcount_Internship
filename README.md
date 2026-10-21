# 📊 PHA Monthly HR Report Automation Suite

> **End-to-end automated ETL pipeline** for monthly HR headcount, demographics, contracts, cost centers, and organizational reporting.
> Replaces error-prone manual Excel VLOOKUPs and joins with a unified, audited Python workflow.

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-Proprietary-red.svg)]()
[![Status](https://img.shields.io/badge/Status-Production%20Ready%20v1.5-brightgreen.svg)]()

---

## 📋 Table of Contents

- [Overview](#-overview)
- [System Architecture](#-system-architecture)
- [Folder & File Structure](#-folder--file-structure)
- [Pipeline Modules Breakdown](#-pipeline-modules-breakdown)
- [Module Settings & Toggle](#-module-settings--toggle-configsettingsjson)
- [Output Column Reference (46 Columns)](#-output-column-reference-all-46-columns)
- [Quick Start Guide](#-quick-start-guide)
- [Distribution Packaging (v1.5)](#-distribution-packaging-v15)
- [Key Design Principles & Data Integrity](#-key-design-principles--data-integrity)
- [Troubleshooting & FAQ](#-troubleshooting--faq)

---

## 🎯 Overview

The **PHA Monthly HR Report Automation Suite** consolidates disparate SAP HR infotypes and master files into a single, standardized, 46-column executive HR master report (`FINAL_HR_Topic_Report_YYYYMMDD.xlsx`).

### ✨ Core Capabilities:
- **Dual Cost Center Mapping:** Independent mapping of Salary Cost Center (`PHA_PA0027_01` + `PHA_CSKT_01`) and General Expense Cost Center (`PHA_PA0027_02` + `PHA_CSKT_02`).
- **Smart Contract Sorting:** Deduplicates multiple historical contract entries (`PHA_PA0016`), sorting descending by `Valid Until` date so the most recent active contract is always captured.
- **Hierarchy & Highest Qualification:** Automatically ranks education records (`Order_PHA_PA0022`) to filter out training entries and select the true highest degree attained.
- **Family Count & Direct Supervisor:** Counts family members (`PHA_PA0021`) and links direct supervisor NIK and Name (`PHA_Direct_Spv`).
- **Normalized Date Formatting:** Enforces unified `DD/MM/YYYY` date representation across Join Dates, Birth Dates, Contract End Dates, and Position Effective Dates.
- **Automated Data Integrity Audit:** Runs post-merge insight validation (`insight_validator.py`) to detect unmatched cost centers, missing supervisor linkages, or invalid dates, exporting an audit log.

---

## 🏗 System Architecture

```
                                  [INPUT DATA]
       +-------------------------------+-------------------------------+
       |                               |                               |
  [Domain Inputs]             [Infotypes & Lookups]             [Shared Files]
  • Contract Type             • PA0016 / PA0022 / PA0027        • PHA_PQAH (Base)
  • Education / Branch        • PA0041 / HRP1001 / PA9000       • PHA_Headings (Template)
  • General Cost Centers      • CSKT_01 / CSKT_02               • PHA_Direct_Spv
  • Position / Tenure         • Branch of Study Lookup          • PHA_PA0021 (Family)
                                                                • PHA_PA0105 (Email)
                                                                • PHA_Job_Layer
                                       │
                                       ▼
                              [ETL MODULE PIPELINE]
  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
  │   Module 1   │    │   Module 2   │    │   Module 3   │    │   Module 4   │
  │ ContractType │    │  Education   │    │EntityDivision│    │  CostCenter  │
  │ PA0016+Lookup│    │ Highest Order│    │PA9000 OrgMap │    │PA0027+CSKT01 │
  │ Date Normal. │    │ + Branch Text│    │ Area/Division│    │PA0027+CSKT02 │
  └──────┬───────┘    └──────┬───────┘    └──────┬───────┘    └──────┬───────┘
         │                   │                   │                   │
         └──────────────┐    │    ┌──────────────┘                   │
                        ▼    ▼    ▼                                  │
  ┌────────────────────────────────────────────────────────┐         │
  │       [Module 5: Join Date] & [Module 6: Pos. Tenure]  │◄────────┘
  │       PA0041 (Type 01) + YoS | HRP1001 + Tenure Calc   │
  └───────────────────────────┬────────────────────────────┘
                              │
                              ▼
                     [FINAL ASSEMBLER]
    • Merges all 6 intermediate outputs to Base PQAH (LEFT JOIN)
    • Integrates Shared files: Direct Spv, Email, Family, Job Layer
    • Orders & standardizes output to exact 46 columns
    • Date formatting to DD/MM/YYYY
    • Generates Data Integrity Insight Report
                              │
                              ▼
                 [FINAL_HR_Topic_Report.xlsx]
                 + 00_processing_summary.xlsx
```

---

## 📁 Folder & File Structure

```
PHA Monthly HR Report Automation/
├── config/
│   └── settings.json                    # Module toggles (enable/disable modules)
├── input/
│   ├── Contract Type & Contract End Data/
│   │   ├── Contract Type_PHA_PA0016.xlsx # Lookup: Contract Type code to description
│   │   └── PHA_PA0016.xlsx              # Employee contract records & Valid Until
│   ├── Education, Institute, Branch of Study/
│   │   ├── Order_PHA_PA0022.xlsx        # Hierarchy ranking for highest qualification
│   │   ├── PHA_Branch of Study.xlsx     # Lookup: Branch code to text
│   │   └── PHA_PA0022.xlsx              # Employee education records
│   ├── Entity, Division, Area, Function/
│   │   └── PHA_PA9000.xlsx              # Custom organizational mapping table
│   ├── General Cost Center & Cost Center Text/
│   │   ├── PHA_PA0027_01.xlsx           # Salary Cost Center assignments
│   │   ├── PHA_CSKT_01.xlsx             # Salary Cost Center descriptions
│   │   ├── PHA_PA0027_02.xlsx           # General Expense Cost Center assignments
│   │   └── PHA_CSKT_02.xlsx             # General Expense Cost Center descriptions
│   ├── Join Date & Year of Service/
│   │   └── PHA_PA0041.xlsx              # Date specifications (Date Type '01' = Join Date)
│   ├── Position Effective Date & Current Tenure/
│   │   └── PHA_HRP1001.xlsx             # Position relationship dates
│   └── Shared/
│       ├── PHA_PQAH.xlsx                # Primary Master Data source (Base population)
│       ├── PHA_Headings.xlsx            # Standard 46-column output header template
│       ├── PHA_Direct_Spv.xlsx          # Superior NIK & Superior Name mapping
│       ├── PHA_Job_Layer.xlsx           # Job ID to Layer hierarchy mapping
│       ├── PHA_PA0021.xlsx              # Family member records (for headcount count)
│       └── PHA_PA0105.xlsx              # Employee communication / corporate emails
├── output/
│   ├── intermediate/                    # Intermediate enriched Excel outputs
│   ├── 00_processing_summary.xlsx       # Status summary of all pipeline stages
│   └── FINAL_HR_Topic_Report_*.xlsx     # Final assembled master report (46 columns)
├── logs/                                # Timestamped run logs & insight audit files
├── src/
│   ├── __init__.py
│   ├── assembler.py                     # Final merger, mapping, and column arranger
│   ├── common.py                        # Shared normalization utilities & date parsers
│   ├── contract_type.py                 # Module 1: Contract type description & dates
│   ├── cost_center.py                   # Module 4: Dual PA0027 + CSKT merger
│   ├── education.py                     # Module 2: Highest degree & branch lookup
│   ├── entity_division.py               # Module 3: Entity, Division, Area, Function
│   ├── excel_export.py                  # Safe Excel writer with lock detection
│   ├── insight_validator.py             # Data integrity auditor and discrepancy tracker
│   ├── join_date.py                     # Module 5: Join Date & Years of Service
│   ├── logger.py                        # Logging configuration
│   ├── position_tenure.py               # Module 6: Position effective date & tenure
│   ├── summary.py                       # Summary table exporter
│   └── validator.py                     # Schema, column, and row-count integrity checks
├── make_package.py                      # Standalone distribution packager (v1.5)
├── requirements.txt                     # Python package dependencies
├── run.py                               # CLI entry point to execute the pipeline
└── RUN_ME.bat                           # 1-click execution batch script for end-users
```

---

## ⚙️ Pipeline Modules Breakdown

### 🔹 Module 1: Contract Type & End Date (`src/contract_type.py`)
- **Inputs:** `PHA_PA0016.xlsx`, `Contract Type_PHA_PA0016.xlsx`
- **Output:** `output/intermediate/PHA_PA0016_With_ContractType_Description.xlsx`
- **Business Logic:** Zero-pads contract type codes (`"2"` ➔ `"02"`) to match lookup table keys. Normalizes `Valid Until` dates to uniform `DD/MM/YYYY` format. In the assembler, rows are sorted descending by date to guarantee the most recent active contract is picked.

### 🔹 Module 2: Education & Branch (`src/education.py`)
- **Inputs:** `PHA_PA0022.xlsx`, `Order_PHA_PA0022.xlsx`, `PHA_Branch of Study.xlsx`
- **Output:** `output/intermediate/PHA_PA0022_HighestOrder_IgnoreTraining_WithBranchText.xlsx`
- **Business Logic:** Filters out non-degree training records, ranks education entries by priority order, picks the top qualification per employee, and joins branch text descriptions.

### 🔹 Module 3: Entity, Division, Area, Function (`src/entity_division.py`)
- **Inputs:** `PHA_PA9000.xlsx`, `PHA_PQAH.xlsx`
- **Output:** `output/intermediate/PHA_PQAH_With_Entity_DivisionDesc_Area_Function.xlsx`
- **Business Logic:** Maps organizational structure (Entity, Division, Area, Function) onto base employee records.

### 🔹 Module 4: General & Salary Cost Center (`src/cost_center.py`)
- **Inputs:** `PHA_PA0027_01.xlsx` + `PHA_CSKT_01.xlsx` (Salary), `PHA_PA0027_02.xlsx` + `PHA_CSKT_02.xlsx` (General Expense)
- **Output:** `output/intermediate/PHA_PA0027_Enriched_CostCenterDesc.xlsx`
- **Business Logic:** Standardizes 10-character cost center IDs and enriches both General Expense and Salary cost centers with their respective official descriptions.

### 🔹 Module 5: Join Date & Years of Service (`src/join_date.py`)
- **Inputs:** `PHA_PA0041.xlsx`, `PHA_PQAH.xlsx`
- **Output:** `output/intermediate/PQAH_Enriched_JoinDate_YoS.xlsx`
- **Business Logic:** Scans date type pairs for Date Type `'01'` (Join Date) and computes total Years of Service.

### 🔹 Module 6: Position Effective Date & Tenure (`src/position_tenure.py`)
- **Inputs:** `PHA_HRP1001.xlsx`, `PHA_PQAH.xlsx`
- **Output:** `output/intermediate/PQAH_Enriched_PositionTenure.xlsx`
- **Business Logic:** Extracts position start dates and computes current tenure in position.

### 🔹 Final Step: Topic Report Assembler (`src/assembler.py`)
- **Inputs:** `PHA_PQAH.xlsx`, `PHA_Headings.xlsx`, all intermediate outputs + Shared files (`PHA_Direct_Spv.xlsx`, `PHA_Job_Layer.xlsx`, `PHA_PA0021.xlsx`, `PHA_PA0105.xlsx`)
- **Output:** `output/FINAL_HR_Topic_Report_YYYYMMDD.xlsx`
- **Business Logic:** Performs left-joins of all domain modules against `_pno_norm` (normalized NIK), counts family members, normalizes all date columns, applies heading ordering, and executes data integrity checks.

---

## 🎛 Module Settings & Toggle (`config/settings.json`)

You can selectively enable or disable individual pipeline modules:

```json
{
    "run_modules": {
        "contract_type":   true,
        "education":       true,
        "entity_division": true,
        "cost_center":     true,
        "join_date":       true,
        "position_tenure": true,
        "assembler":       true
    }
}
```

---

## 📊 Output Column Reference (All 46 Columns)

| # | Target Column Name | Primary Source / Extraction Logic |
|:---:|:---|:---|
| 1 | **Entity** | `PHA_PA9000.xlsx` via `PHA_PQAH.xlsx` |
| 2 | **Division** | `PHA_PA9000.xlsx` |
| 3 | **Area** | `PHA_PA9000.xlsx` |
| 4 | **Function** | `PHA_PA9000.xlsx` |
| 5 | **Personnel No.** | `PHA_PQAH.xlsx` (Employee NIK) |
| 6 | **Personnel Number** | `PHA_PQAH.xlsx` (Employee Full Name) |
| 7 | **Lvl** | `PHA_PQAH.xlsx` (Employee Level) |
| 8 | **CoCd** | `PHA_PQAH.xlsx` (Company Code) |
| 9 | **Company Name** | `PHA_PQAH.xlsx` |
| 10 | **ESgrp** | `PHA_PQAH.xlsx` (Employee Subgroup Code) |
| 11 | **Name of EE Subgroup** | `PHA_PQAH.xlsx` |
| 12 | **PArea** | `PHA_PQAH.xlsx` (Personnel Area Code) |
| 13 | **Personnel Area Text** | `PHA_PQAH.xlsx` |
| 14 | **Subarea** | `PHA_PQAH.xlsx` (Personnel Subarea Code) |
| 15 | **P. SubArea Text** | `PHA_PQAH.xlsx` |
| 16 | **Position** | `PHA_PQAH.xlsx` (Position ID) |
| 17 | **Position Name** | `PHA_PQAH.xlsx` |
| 18 | **Org. Unit** | `PHA_PQAH.xlsx` (Organizational Unit ID) |
| 19 | **Name of Organizational Unit** | `PHA_PQAH.xlsx` |
| 20 | **Join Date** | `PHA_PA0041.xlsx` (Date Type `'01'`) |
| 21 | **Year of Service** | Calculated: Join Date ➔ Today (Years) |
| 22 | **Job** | `PHA_PQAH.xlsx` (Job Code) |
| 23 | **Job Title** | `PHA_PQAH.xlsx` |
| 24 | **Gender** | `PHA_PQAH.xlsx` |
| 25 | **Birth Date** | `PHA_PQAH.xlsx` |
| 26 | **Age** | Calculated: Birth Date ➔ Today (Years) |
| 27 | **Nationality** | `PHA_PQAH.xlsx` |
| 28 | **Birthplace** | `PHA_PQAH.xlsx` |
| 29 | **Education** | `PHA_PA0022.xlsx` + `Order_PHA_PA0022.xlsx` (Highest degree) |
| 30 | **Institute** | `PHA_PA0022.xlsx` |
| 31 | **Branch Study** | `PHA_PA0022.xlsx` + `PHA_Branch of Study.xlsx` |
| 32 | **Religion** | `PHA_PQAH.xlsx` |
| 33 | **Number of Family Members** | `PHA_PA0021.xlsx` (Count of records per employee) |
| 34 | **Marital Status** | `PHA_PQAH.xlsx` |
| 35 | **Layer** | `PHA_Job_Layer.xlsx` |
| 36 | **Superior NIK** | `PHA_Direct_Spv.xlsx` (`Personnel` ➔ `D.Superior`) |
| 37 | **Superior Name** | `PHA_Direct_Spv.xlsx` (`Direct Sup`) |
| 38 | **Salary Cost Center (S4)** | `PHA_PA0027_01.xlsx` |
| 39 | **Salary Cost Center Desc (S4)** | `PHA_CSKT_01.xlsx` |
| 40 | **General Expense Cost Center** | `PHA_PA0027_02.xlsx` |
| 41 | **General Expense Cost Center Desc** | `PHA_CSKT_02.xlsx` |
| 42 | **Email** | `PHA_PA0105.xlsx` (Corporate Email) |
| 43 | **Contract Type** | `PHA_PA0016.xlsx` + `Contract Type_PHA_PA0016.xlsx` |
| 44 | **Contract End Date** | `PHA_PA0016.xlsx` (Sorted latest `Valid Until`) |
| 45 | **Position Effective Date** | `PHA_HRP1001.xlsx` |
| 46 | **Current Tenure** | Calculated: Position Effective Date ➔ Today (Years) |

---

## 🚀 Quick Start Guide

### Prerequisites
- Python 3.10+
- Microsoft Excel installed (optional, for viewing reports)

### 1. Installation
```powershell
# Clone the repository
git clone https://github.com/Lusive12/ETL_MergeExcelHeadings_Internship.git
cd "PHA Monthly HR Report Automation"

# Create and activate virtual environment
python -m venv .venv
.venv\Scriptsctivate

# Install dependencies
pip install -r requirements.txt
```

### 2. Prepare Input Data
Place the monthly SAP export files into their respective subfolders in `input/` according to the structure above.

### 3. Run the Automation
- **Option A (One-Click for Users):** Double-click `RUN_ME.bat`.
- **Option B (Command Line):**
  ```powershell
  python run.py
  ```

### 4. Review Results
- The final topic report will be saved at: `output/FINAL_HR_Topic_Report_YYYYMMDD.xlsx`.
- The execution summary is recorded at: `output/00_processing_summary.xlsx`.
- Detailed logs and data integrity insight reports are saved in `logs/`.

---

## 📦 Distribution Packaging (v1.5)

To generate a standalone distribution package for end users without requiring Git:

```powershell
python make_package.py
```

This compiles a clean, production-ready package under:
`dist/PHA_Monthly_HR_Automation_Packagev1.5/`

---

## 🛡 Key Design Principles & Data Integrity

- **Strict Key Normalization:** All join keys (`_pno_norm`, Cost Centers, Contract Type Codes) are sanitized, stripping whitespace, removing floating-point artifacts (`.0`), and zero-padded where appropriate.
- **Row Count Preservation:** Merges enforce left-joins to guarantee that no base employees from `PHA_PQAH.xlsx` are dropped.
- **Excel Lock Protection:** `excel_export.py` checks file access prior to writing, warning users clearly if an output file is currently open in Excel.
- **Automated Anomaly Audit:** `insight_validator.py` flags missing direct supervisors, unmatched cost centers, and unpopulated fields for review.

---

## 💬 Troubleshooting & FAQ

| Issue | Cause | Solution |
|---|---|---|
| **"Cannot overwrite... file is open in Excel"** | The output Excel report is open in Microsoft Excel. | Close Excel and re-run `run.py` or `RUN_ME.bat`. |
| **"File is not a zip file"** | The input file is corrupt or saved in an older non-standard `.xls` or HTML format. | Re-export the file as standard `.xlsx` from SAP. |
| **"Missing required column"** | Input file headers do not match the expected column naming. | Check the input file headers against the specification in this README. |

---

*PHA Monthly HR Report Automation Suite — Engineered for high precision, data integrity, and complete reporting autonomy.*
