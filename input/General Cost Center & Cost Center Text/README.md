# 📂 General Cost Center & Cost Center Text

**Module 4** — Enriches each employee's IT0027 cost center assignment with the cost center description text.

---

## Files

### `PHA_PA0027_01.xlsx` — Salary Cost Center Assignments (SAP IT0027)
- **Export source:** SAP Infotype PA0027
- **Key column:** `Personnel number`
- **Contains:** Multiple cost center columns per record (salary split)
- An employee may have multiple PA0027 records; the module picks the **most valid** one based on date ranges.

### `PHA_PA0027_02.xlsx` — General Expense Cost Center Assignments (SAP IT0027)
- **Export source:** SAP Infotype PA0027
- **Key column:** `Personnel number`
- **Contains:** Multiple cost center columns per record (general expense split)
- An employee may have multiple PA0027 records; the module picks the **most valid** one based on date ranges.

### `PHA_CSKT_01.xlsx` — Salary Cost Center Master Lookup
- **Role:** Maps cost center codes to their descriptions
- **Key column:** `Cost Center` (or equivalent)
- **Value column:** Cost Center description text
- **Selection rule:** For employees with multiple date-range records, the entry with the **maximum Valid To date** is selected.

### `PHA_CSKT_02.xlsx` — General Expense Cost Center Master Lookup
- **Role:** Maps cost center codes to their descriptions
- **Key column:** `Cost Center` (or equivalent)
- **Value column:** Cost Center description text
- **Selection rule:** For employees with multiple date-range records, the entry with the **maximum Valid To date** is selected.

---

## Output

`output/intermediate/IT0027_Enriched_CostCenterDesc.xlsx`

Added columns:
- **`Salary Cost Center (S4)`** — Salary cost center code
- **`Salary Cost Center Desc (S4)`** — Salary cost center description
- **`General Expense Cost Center`** — General expense cost center code
- **`General Expense Cost Center Desc`** — General expense cost center description
