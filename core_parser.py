# core_parser.py
"""
核心解析引擎：100% 完整保留台大牙周專科病歷格式生成演算法，確保數據鏡射與過濾精準。
"""
import pandas as pd
import io
from typing import Dict, List, Set, Any, Tuple

def clean_cell(x): 
    return str(x).strip() if pd.notna(x) else ""

def is_valid_tooth(x):
    x = clean_cell(x)
    if not x.isdigit(): return False
    tooth = int(x)
    return (tooth // 10) in [1, 2, 3, 4] and 1 <= (tooth % 10) <= 8

def is_comparison_file(df) -> bool:
    full_text = df.astype(str).to_string()
    return ("Date(Re-evaluation)" in full_text) or ("Re=" in full_text) or ("I" in df.values and "R" in df.values)

def find_tooth_rows(df):
    tooth_rows = []
    for i in range(len(df)):
        row_cells = [clean_cell(x) for x in df.iloc[i].tolist()]
        valid_teeth = [int(x) for x in row_cells if is_valid_tooth(x)]
        if "Tooth" in row_cells or len(valid_teeth) >= 5:
            tooth_rows.append(i)
    return tooth_rows

def find_missing_rows(df):
    return [i for i in range(len(df)) if "MISSING" in [clean_cell(x).upper() for x in df.iloc[i].tolist()]]

def find_furcation_rows(df):
    furcation_rows = []
    for i in range(len(df)):
        row_text = " ".join([clean_cell(x) for x in df.iloc[i].tolist()]).lower()
        if "furcation" in row_text:
            furcation_rows.append({"header_row": i, "label_row": i, "value_row": i + 1})
    return furcation_rows

def get_three_digit_raw_list(df, row_idx, start_col):
    if row_idx is None or pd.isna(row_idx): return ["?", "?", "?"]
    values = []
    for c in range(start_col, start_col + 3):
        if c < df.shape[1]:
            v = clean_cell(df.iloc[int(row_idx), c])
            if v == "": v = "?"
            values.append(v)
        else:
            values.append("?")
    return values

def get_tooth_start_columns(df, tooth_row_idx):
    return {int(clean_cell(df.iloc[tooth_row_idx, col])): col for col in range(df.shape[1]) if is_valid_tooth(df.iloc[tooth_row_idx, col])}

def get_missing_teeth_set(df, tooth_rows, missing_rows):
    missing_teeth = set()
    if len(tooth_rows) < 2: return missing_teeth
    up_m = missing_rows[0] if len(missing_rows) > 0 else None
    lo_m = missing_rows[-1] if len(missing_rows) > 1 else (missing_rows[0] if len(missing_rows) == 1 else None)
    for t, c in get_tooth_start_columns(df, tooth_rows[0]).items():
        if t // 10 in [1, 2] and up_m is not None and c + 1 < df.shape[1] and clean_cell(df.iloc[up_m, c + 1]).upper() == "TRUE": missing_teeth.add(t)
    for t, c in get_tooth_start_columns(df, tooth_rows[-1]).items():
        if t // 10 in [3, 4] and lo_m is not None and c + 1 < df.shape[1] and clean_cell(df.iloc[lo_m, c + 1]).upper() == "TRUE": missing_teeth.add(t)
    return missing_teeth

def collect_absolute_row_indices(df, target_stage="I"):
    """智慧型列號收集器：極致相容 Initial (I) 與 Re-evaluation (R)"""
    is_comp = is_comparison_file(df)
    pd_rows, gm_rows, mob_rows, km_rows = [], [], [], []

    for i in range(len(df)):
        c0 = clean_cell(df.iloc[i, 0]).upper()
        c1 = clean_cell(df.iloc[i, 1]).upper()
        prefix = f"{c0} {c1}"

        if "PD" in prefix:
            r_idx = i if (not is_comp or target_stage == "I") else i + 1
            pd_rows.append(r_idx)
        elif "GM" in prefix or "CEJ" in prefix:
            r_idx = i if (not is_comp or target_stage == "I") else i + 1
            gm_rows.append(r_idx)
        elif "MOBILITY" in prefix and "SCALE" not in prefix:
            r_idx = i if (not is_comp or target_stage == "I") else i + 1
            mob_rows.append(r_idx)
        elif "KM" in prefix:
            r_idx = i if (not is_comp or target_stage == "I") else i + 1
            km_rows.append(r_idx)

    return {
        "up_b_pd": pd_rows[0] if len(pd_rows) >= 1 else None, 
        "up_p_pd": pd_rows[1] if len(pd_rows) >= 2 else None,
        "lo_l_pd": pd_rows[2] if len(pd_rows) >= 3 else None, 
        "lo_b_pd": pd_rows[3] if len(pd_rows) >= 4 else None,
        "up_b_gm": gm_rows[0] if len(gm_rows) >= 1 else None, 
        "up_p_gm": gm_rows[1] if len(gm_rows) >= 2 else None,
        "lo_l_gm": gm_rows[2] if len(gm_rows) >= 3 else None, 
        "lo_b_gm": gm_rows[3] if len(gm_rows) >= 4 else None,
        "up_mob": mob_rows[0] if len(mob_rows) >= 1 else None, 
        "lo_mob": mob_rows[-1] if len(mob_rows) >= 1 else None,
        "up_km":  km_rows[0] if len(km_rows) >= 1 else None, 
        "lo_km":  km_rows[-1] if len(km_rows) >= 1 else None
    }

def parse_periodontal_csv(file_stream) -> Tuple[Any, Set[int], bool, Dict[str, Any]]:
    try:
        file_stream.seek(0)
        df = pd.read_csv(file_stream, header=None)
    except Exception:
        file_stream.seek(0)
        df = pd.read_excel(file_stream, header=None)

    is_comp = is_comparison_file(df)
    patient_info = {}

    for r in range(min(10, len(df))):
        row = df.iloc[r].dropna().tolist()
        for idx, val in enumerate(row):
            val_str = str(val).strip()
            if "Patient's Name" in val_str and idx + 2 < len(row):
                patient_info["name"] = str(row[idx + 2]).strip()
            elif "Case Report No." in val_str and idx + 2 < len(row):
                patient_info["case_no"] = str(row[idx + 2]).strip()

    tooth_rows = find_tooth_rows(df)
    missing_rows = find_missing_rows(df)
    missing_teeth = get_missing_teeth_set(df, tooth_rows, missing_rows)

    return df, missing_teeth, is_comp, patient_info
