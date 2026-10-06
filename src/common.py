"""
common.py - Shared utilities for all HR automation modules.

Implements the 5 Common Business Rules from the specification:
  Rule 1  Normalize numeric IDs  (2102.0 -> "2102")
  Rule 2  Trim spaces            ("  NIK  " -> "NIK")
  Rule 3  Key lookup as string   (all merge keys cast to str)
  Rule 4  Row preservation       (all merges must be LEFT JOIN)
  Rule 5  Auditability           (matched / unmatched / blank counts)
"""
from pathlib import Path
from typing import Optional
import pandas as pd


def normalize_id(series: pd.Series) -> pd.Series:
    """
    Rule 1 + 2: Convert to string, strip whitespace, remove trailing .0
    """
    return (
        series.fillna("")
              .astype(str)
              .str.strip()
              .str.replace(r"\.0$", "", regex=True)
    )


def normalize_date_type(series: pd.Series) -> pd.Series:
    """
    Normalize SAP date type codes to zero-padded 2-character strings.
    """
    s = series.fillna("").astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    return s.apply(lambda x: x.zfill(2) if x != "" else "")


def strip_leading_zeros(series: pd.Series) -> pd.Series:
    """
    Strip leading zeros from string series, keeping at least one character.
    """
    return series.apply(lambda x: x.lstrip("0") or "0" if isinstance(x, str) else x)


def load_excel(path: Path) -> pd.DataFrame:
    """
    Load an Excel file, reading all columns as strings.
    Strips whitespace from all column names.
    """
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_column_ci(df: pd.DataFrame, target_name: str) -> Optional[str]:
    """
    Find exact column name in DataFrame using exact case-insensitive & trimmed comparison.
    No alias substitution - matches target_name ignoring case and whitespace only.
    """
    target_norm = target_name.strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == target_norm:
            return str(col)
    return None


def get_required_column_ci(df: pd.DataFrame, target_name: str) -> str:
    """
    Get required column name in DataFrame case-insensitively, or raise ValueError.
    """
    col = find_column_ci(df, target_name)
    if col is None:
        raise ValueError(f"Missing required column: '{target_name}' (checked case-insensitively)")
    return col


def detect_date_format(series: pd.Series) -> str:
    """
    Determine the date format of an entire column by scanning for unambiguous values.

    Strategy (column-level, not per-cell):
      - Scan all non-null values in the column.
      - If the YEAR is first (>1000) -> ISO format (YYYY-MM-DD).
      - If the FIRST part >12 -> day is first -> format is DD/MM/YYYY.
      - If the SECOND part >12 -> second part can't be a month -> format is MM/DD/YYYY.
      - If no unambiguous value found -> assume SAP default: MM/DD/YYYY.

    Returns one of: "ISO", "DD/MM/YYYY", "MM/DD/YYYY"
    """
    import re
    for raw in series.dropna():
        v = str(raw).strip()
        date_part = re.split(r"\s", v)[0]
        parts = re.split(r"[/\-]", date_part)
        if len(parts) != 3:
            continue
        try:
            a, b, c = int(parts[0]), int(parts[1]), int(parts[2])
        except ValueError:
            continue
        if a > 1000:
            return "ISO"
        if c > 1000:
            if a > 12:
                return "DD/MM/YYYY"
            if b > 12:
                return "MM/DD/YYYY"
    return "MM/DD/YYYY"


def normalize_date_column(series: pd.Series) -> pd.Series:
    """
    Normalize an entire date column to DD/MM/YYYY format.
    """
    fmt = detect_date_format(series)

    if fmt == "ISO":
        parse_fmt = "%Y-%m-%d %H:%M:%S"
        fallback_fmt = "%Y-%m-%d"
    elif fmt == "MM/DD/YYYY":
        parse_fmt = "%m/%d/%Y"
        fallback_fmt = None
    else:
        parse_fmt = "%d/%m/%Y"
        fallback_fmt = None

    def _convert(val):
        if not isinstance(val, str) or val.strip() == "":
            return val
        v = val.strip()
        try:
            return pd.to_datetime(v, format=parse_fmt).strftime("%d/%m/%Y")
        except Exception:
            if fallback_fmt:
                try:
                    return pd.to_datetime(v, format=fallback_fmt).strftime("%d/%m/%Y")
                except Exception:
                    pass
        return val

    return series.apply(_convert)


def audit_counts(df: pd.DataFrame, lookup_col: str) -> dict:
    """
    Rule 5: Return matched / unmatched / blank counts for a lookup result column.
    """
    total     = len(df)
    blank     = df[lookup_col].isna() | (df[lookup_col].astype(str).str.strip() == "")
    matched   = int((~blank).sum())
    unmatched = int(blank.sum())
    return {"total": total, "matched": matched, "unmatched": unmatched}