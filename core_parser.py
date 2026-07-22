# core_parser.py
"""
核心解析引擎：100% 完整保留台大牙周專科病歷格式生成演算法，確保數據鏡射與過濾精準。
支援 Initial 與 Initial & Re-evaluation 雙期分側定位，修正 CAL 被 MOBILITY SCALE 誤抓的問題。
"""
import pandas as pd
import io
from typing import Dict, List, Set, Any, Tuple

def clean_cell(x): 
    if pd.isna(x):
        return ""
    val_str = str(x).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

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
        row_text = " ".join([clean_cell(df.iloc[i, c]) for c in range(df.shape[1])]).lower()
        # 🚀 排除 "grade" 與 "scale" 等說明標題列，精確鎖定真正的 Furcation 數據標籤列
        if "furcation" in row_text and "grade" not in row_text and "scale" not in row_text:
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

def collect_comparison_row_indices(df):
    """智慧型分區列號收集器：修復 CAL 被 MOBILITY SCALE 誤讀之 Bug"""
    is_comp = is_comparison_file(df)

    midpoint = len(df) // 2
    for r in range(len(df)):
        row_str = " ".join([str(df.iloc[r, c]) for c in range(df.shape[1]) if pd.notna(df.iloc[r, c])])
        if "48" in row_str and "38" in row_str:
            midpoint = r
            break

    res = {}

    # 1. 上顎 (Upper Arch: 0 ~ midpoint)
    for r in range(0, midpoint):
        c0 = clean_cell(df.iloc[r, 0]).upper()
        c1 = clean_cell(df.iloc[r, 1]).upper()
        prefix = f"{c0} {c1}"

        if "PD" in prefix and "up_b_pd_i" not in res:
            res["up_b_pd_i"] = r; res["up_b_pd_r"] = r + 1 if is_comp else r
        elif "PD" in prefix and "up_b_pd_i" in res and "up_p_pd_i" not in res:
            res["up_p_pd_i"] = r; res["up_p_pd_r"] = r + 1 if is_comp else r

        if ("GM" in prefix or "RECESSION" in prefix or "CEJ" in prefix) and "up_b_gm_i" not in res:
            res["up_b_gm_i"] = r; res["up_b_gm_r"] = r + 1 if is_comp else r
        elif ("GM" in prefix or "RECESSION" in prefix or "CEJ" in prefix) and "up_b_gm_i" in res and "up_p_gm_i" not in res:
            res["up_p_gm_i"] = r; res["up_p_gm_r"] = r + 1 if is_comp else r

        # 🚀 排除 SCALE，精確匹配 CAL
        if "CAL" in prefix and "SCALE" not in prefix and "up_b_cal_i" not in res:
            res["up_b_cal_i"] = r; res["up_b_cal_r"] = r + 1 if is_comp else r
        elif "CAL" in prefix and "SCALE" not in prefix and "up_b_cal_i" in res and "up_p_cal_i" not in res:
            res["up_p_cal_i"] = r; res["up_p_cal_r"] = r + 1 if is_comp else r

        if "KM" in prefix and "up_km_i" not in res:
            res["up_km_i"] = r; res["up_km_r"] = r + 1 if is_comp else r

        if "MOBILITY" in prefix and "SCALE" not in prefix and "up_mob_i" not in res:
            res["up_mob_i"] = r; res["up_mob_r"] = r + 1 if is_comp else r

    # 2. 下顎 (Lower Arch: midpoint ~ len(df))
    for r in range(midpoint, len(df)):
        c0 = clean_cell(df.iloc[r, 0]).upper()
        c1 = clean_cell(df.iloc[r, 1]).upper()
        prefix = f"{c0} {c1}"

        if "PD" in prefix and "lo_l_pd_i" not in res:
            res["lo_l_pd_i"] = r; res["lo_l_pd_r"] = r + 1 if is_comp else r
        elif "PD" in prefix and "lo_l_pd_i" in res and "lo_b_pd_i" not in res:
            res["lo_b_pd_i"] = r; res["lo_b_pd_r"] = r + 1 if is_comp else r

        if ("GM" in prefix or "RECESSION" in prefix or "CEJ" in prefix) and "lo_l_gm_i" not in res:
            res["lo_l_gm_i"] = r; res["lo_l_gm_r"] = r + 1 if is_comp else r
        elif ("GM" in prefix or "RECESSION" in prefix or "CEJ" in prefix) and "lo_l_gm_i" in res and "lo_b_gm_i" not in res:
            res["lo_b_gm_i"] = r; res["lo_b_gm_r"] = r + 1 if is_comp else r

        # 🚀 排除 SCALE，精確匹配 CAL
        if "CAL" in prefix and "SCALE" not in prefix and "lo_l_cal_i" not in res:
            res["lo_l_cal_i"] = r; res["lo_l_cal_r"] = r + 1 if is_comp else r
        elif "CAL" in prefix and "SCALE" not in prefix and "lo_l_cal_i" in res and "lo_b_cal_i" not in res:
            res["lo_b_cal_i"] = r; res["lo_b_cal_r"] = r + 1 if is_comp else r

        if "KM" in prefix and "lo_km_i" not in res:
            res["lo_km_i"] = r; res["lo_km_r"] = r + 1 if is_comp else r

        if "MOBILITY" in prefix and "SCALE" not in prefix and "lo_mob_i" not in res:
            res["lo_mob_i"] = r; res["lo_mob_r"] = r + 1 if is_comp else r

    return res

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
                patient_info["name"] = clean_cell(row[idx + 2])
            elif "Case Report No." in val_str and idx + 2 < len(row):
                patient_info["case_no"] = clean_cell(row[idx + 2])

    tooth_rows = find_tooth_rows(df)
    missing_rows = find_missing_rows(df)
    missing_teeth = get_missing_teeth_set(df, tooth_rows, missing_rows)

    return df, missing_teeth, is_comp, patient_info
