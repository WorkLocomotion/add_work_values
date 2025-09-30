#!/usr/bin/env python3
# Enrich matched company job titles with O*NET Work Values scores.

import argparse
import io
import os
import re
import sys
from typing import List, Optional

import pandas as pd
import warnings

try:
    import requests
except ImportError:
    requests = None  # Will be available if installed via requirements

# Canonical Work Values order
VALUE_COLS_CANON = [
    "Achievement",
    "Independence",
    "Recognition",
    "Relationships",
    "Support",
    "Working Conditions",
]

# Acceptable header variants for each Work Value (canonical -> acceptable tokens)
VALUE_ALIASES = {
    "Achievement": ["achievement"],
    "Independence": ["independence"],
    "Recognition": ["recognition"],
    "Relationships": ["relationship", "relationships"],
    "Support": ["support"],
    "Working Conditions": ["workingconditions", "workconditions", "workingcondition", "workcondition"],
}

# --------- Helpers ---------
def _canon(s: str) -> str:
    """Lowercase and strip non-alphanumerics for tolerant header matching."""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def norm_soc(s: str) -> str:
    """Normalize SOC code best-effort (e.g., '131081' -> '13-1081')."""
    if not isinstance(s, str):
        s = str(s) if pd.notna(s) else ""
    s = s.strip()
    if not s:
        return ""
    digits = re.sub(r"[^0-9]", "", s)
    if len(digits) >= 7:
        left = f"{digits[0:2]}-{digits[2:7]}"
        rest = digits[7:]
        return f"{left}{('-' + rest) if rest else ''}"
    elif len(digits) in (5, 6):
        return f"{digits[:2]}-{digits[2:]}"
    else:
        return s

def find_col(df: pd.DataFrame, options: List[str]) -> str:
    """Find a column by canonicalized name (case/punct/space-insensitive)."""
    want = {_canon(o) for o in options}
    for c in df.columns:
        if _canon(c) in want:
            return c
    raise KeyError(f"Required column not found. Tried: {options}")

def _col_like(df: pd.DataFrame, options: List[str]) -> Optional[str]:
    """Return the first matching column name from options (canonicalized), else None."""
    want = {_canon(o) for o in options}
    for c in df.columns:
        if _canon(c) in want:
            return c
    return None

# --------- O*NET loader (supports WIDE and LONG) ---------
def load_onet_values(source: str) -> pd.DataFrame:
    """
    Load O*NET Work Values from a local path or GitHub RAW URL.

    Supports:
      A) WIDE format: columns per value (Achievement, Independence, ...)
      B) LONG format: rows per (SOC, Element Name, Scale, Date, Data Value)

    Returns wide DataFrame:
      ['SOC Code', 'Achievement', 'Independence', 'Recognition',
       'Relationships', 'Support', 'Working Conditions']
    """
    # Read file
    if source.lower().startswith("http"):
        if requests is None:
            raise RuntimeError("requests not installed. Please `pip install requests`.")
        resp = requests.get(source, timeout=60)
        resp.raise_for_status()
        data = io.BytesIO(resp.content)
        onet_df = pd.read_excel(data)
    else:
        onet_df = pd.read_excel(source)

    onet_df.columns = [str(c).strip() for c in onet_df.columns]

    # Identify SOC column (tolerant)
    soc_col = None
    soc_opts = [
        "SOC Code", "SOC", "Code", "SOC_Code", "Occupation Code",
        "O*NET-SOC Code", "O*NET-SOC Codes", "ONET SOC Code", "ONET SOC Codes"
    ]
    soc_col = _col_like(onet_df, soc_opts)
    if soc_col is None:
        # heuristic: contains both 'soc' and 'code'
        for c in onet_df.columns:
            cl = _canon(c)
            if "soc" in cl and "code" in cl:
                soc_col = c
                break
    if soc_col is None:
        raise KeyError("Could not find a SOC code column in the O*NET file.")

    # Try WIDE format first (value columns present directly)
    try:
        canon_headers = {_canon(c): c for c in onet_df.columns}

        def _find_value_col(aliases: List[str]) -> Optional[str]:
            # exact canonical match
            for a in aliases:
                if a in canon_headers:
                    return canon_headers[a]
            # contains match (e.g., 'Achievement (Score)')
            for a in aliases:
                for k, orig in canon_headers.items():
                    if a in k:
                        return orig
            return None

        value_cols = []
        for v in VALUE_COLS_CANON:
            aliases = [_canon(x) for x in VALUE_ALIASES.get(v, [_canon(v)])]
            hit = _find_value_col(aliases)
            if hit is None:
                raise KeyError(f"wide-missing:{v}")
            value_cols.append(hit)

        # If all six found => WIDE
        keep = [soc_col] + value_cols
        out = onet_df[keep].copy()
        out.rename(columns={soc_col: "SOC Code"}, inplace=True)
        out["SOC Code"] = out["SOC Code"].apply(norm_soc)
        out.columns = ["SOC Code"] + VALUE_COLS_CANON
        out = out.drop_duplicates(subset=["SOC Code"], keep="first")
        return out

    except KeyError:
        # Fall through to LONG handling
        pass

    # LONG format (Element Name, Data Value)
    # Required columns
    if _col_like(onet_df, ["Element Name"]) is None or _col_like(onet_df, ["Data Value"]) is None:
        raise KeyError(
            "O*NET file is neither recognized wide nor long format. "
            f"Available columns: {list(onet_df.columns)}"
        )

    # Standardize working copy
    df = onet_df.copy()
    ren = {}
    c_elem = _col_like(df, ["Element Name"]);  ren[c_elem] = "Element Name"
    c_val  = _col_like(df, ["Data Value"]);    ren[c_val]  = "Data Value"
    c_scale= _col_like(df, ["Scale Name","Scale"])
    if c_scale: ren[c_scale] = "Scale Name"
    c_date = _col_like(df, ["Date"])
    if c_date: ren[c_date] = "Date"
    if soc_col != "SOC Code": ren[soc_col] = "SOC Code"
    df.rename(columns=ren, inplace=True)

    # Normalize SOC
    df["SOC Code"] = df["SOC Code"].apply(norm_soc)

    # Map element names to canonical values
    df["_elem_canon"] = df["Element Name"].map(_canon)
    alias_to_value = {}
    for canon_value, aliases in VALUE_ALIASES.items():
        for a in aliases:
            alias_to_value[_canon(a)] = canon_value
        alias_to_value[_canon(canon_value)] = canon_value

    df["_target_value"] = df["_elem_canon"].map(alias_to_value)
    df = df[~df["_target_value"].isna()].copy()

    if df.empty:
        raise KeyError(
            "Could not find any of the six Work Values in 'Element Name'. "
            f"Saw examples: {sorted(set(onet_df['Element Name']))[:10]}"
        )

    # Prefer 'importance' scale if available
    if "Scale Name" in df.columns:
        mask_importance = df["Scale Name"].astype(str).str.lower().str.contains("importance")
        if mask_importance.any():
            df = df[mask_importance].copy()

    # If multiple rows per (SOC, value), pick latest Date if present; else last occurrence
    if "Date" in df.columns:
        # coerce to datetime, silence format warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df["__date"] = pd.to_datetime(df["Date"], errors="coerce")
        df.sort_values(["SOC Code", "_target_value", "__date"], ascending=[True, True, True], inplace=True)
        df = df.groupby(["SOC Code", "_target_value"], as_index=False).tail(1)

    wide = df.pivot_table(
        index="SOC Code",
        columns="_target_value",
        values="Data Value",
        aggfunc="last"
    ).reset_index()

    # Ensure all six columns exist
    for v in VALUE_COLS_CANON:
        if v not in wide.columns:
            wide[v] = pd.NA

    wide = wide[["SOC Code"] + VALUE_COLS_CANON].drop_duplicates(subset=["SOC Code"])
    return wide

# --------- Main ---------
def main(argv: List[str] = None) -> int:
    p = argparse.ArgumentParser(description="Enrich matched job titles with O*NET Work Values")
    p.add_argument("--input_excel", required=True, help="Path to 'Company Job Titles - Mapped.xlsx' (or similar)")
    p.add_argument("--onet_values_url", required=True, help="GitHub RAW URL or local path to 'Work Values.xlsx'")
    p.add_argument("--output_excel", default="Company Job Titles - Mapped.with_work_values.xlsx", help="Output Excel path")
    args = p.parse_args(argv)

    # Load input
    try:
        in_df = pd.read_excel(args.input_excel)
    except Exception as e:
        print(f"ERROR: Failed to read input Excel: {e}")
        return 2

    in_df.columns = [str(c).strip() for c in in_df.columns]

    # Columns (tolerant names)
    try:
        job_col = find_col(in_df, ["Job Title", "Job Titles"])
    except KeyError:
        job_col = None

    try:
        soc_col_in = find_col(
            in_df,
            [
                "SOC Code", "SOC", "Code", "SOC_Code",
                "O*NET-SOC Codes", "O*NET-SOC Code",
                "ONET SOC Codes", "ONET SOC Code",
                "ONET-SOC Codes", "ONET-SOC Code",
                "ONET Codes", "ONET Code",
            ],
        )
    except KeyError:
        print("ERROR: Input file must include a 'SOC Code' column.")
        return 3

    try:
        headcount_col = find_col(in_df, ["HeadCount", "Head Count", "HC"])
    except KeyError:
        headcount_col = None

    try:
        occ_title_col = find_col(in_df, ["Occupational Title", "Occupation Title", "ONET Title", "Title"])
    except KeyError:
        occ_title_col = None

    # Normalize/standardize input columns
    in_df["SOC Code"] = in_df[soc_col_in].apply(norm_soc)

    if headcount_col is None:
        in_df["HeadCount"] = 1
    else:
        in_df.rename(columns={headcount_col: "HeadCount"}, inplace=True)

    if occ_title_col is not None:
        in_df.rename(columns={occ_title_col: "Occupational Title"}, inplace=True)
    else:
        in_df["Occupational Title"] = "" if job_col is None else in_df[job_col]

    if job_col is not None and job_col != "Job Title":
        in_df.rename(columns={job_col: "Job Title"}, inplace=True)

    # Load O*NET values (wide/long supported)
    try:
        onet_df = load_onet_values(args.onet_values_url)
    except Exception as e:
        print(f"ERROR: Failed to load O*NET Work Values: {e}")
        return 4

    # Merge and order columns
    merged = pd.merge(in_df, onet_df, on="SOC Code", how="left", validate="m:1", suffixes=("", "_ONET"))

    front = [c for c in ["Job Title", "HeadCount", "SOC Code", "Occupational Title"] if c in merged.columns]
    vals = [c for c in VALUE_COLS_CANON if c in merged.columns]
    others = [c for c in merged.columns if c not in front + vals]
    merged = merged[front + vals + others]

    # Save (auto-increment on permission error)
    out_path = args.output_excel
    base, ext = os.path.splitext(out_path)
    attempt = 0
    while True:
        try:
            merged.to_excel(out_path, index=False)
            break
        except PermissionError:
            attempt += 1
            out_path = f"{base} ({attempt}){ext}"
            if attempt > 10:
                print("ERROR: Could not write output after multiple attempts (file locked?).")
                return 5

    print(f"Saved: {out_path}")
    print("")
    print("WORK LOCOMOTION: Make Potential Actual")
    print("")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())


