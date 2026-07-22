# core_parser.py
"""
核心解析引擎：100% 完整保留台大牙周專科病歷格式生成演算法，確保數據鏡射與過濾精準。
"""
import pandas as pd
import io
from contextlib import redirect_stdout
from typing import Dict, List, Set, Any, Tuple

def clean_cell(x): 
    return str(x).strip()

def is_valid_tooth(x):
    x = clean_cell(x)
    if not x.isdigit(): return False
    tooth = int(x)
    return (tooth // 10) in [1, 2, 3, 4] and 1 <= (tooth % 10) <= 8

def is_comparison_file(df) -> bool:
    """自動判斷上傳的檔案是否為 Initial & Re-evaluation 雙期對比檔"""
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

def find_pd_rows(df):
    return [i for i in range(len(df)) if any(clean_cell(x).upper().startswith("PD") for x in df.iloc[i].tolist())]

def find_mobility_rows(df):
    return [i for i in range(len(df)) if any("Mobility" in clean_cell(x) for x in df.iloc[i].tolist())]

def find_furcation_rows(df):
    furcation_rows = []
    for i in range(len(df)):
        row_text = " ".join([clean_cell(x) for x in df.iloc[i].tolist()]).lower()
        if "furcation" in row_text:
            furcation_rows.append({"header_row": i, "label_row": i, "value_row": i + 1})
    return furcation_rows

def get_three_digit_pd(df, row_idx, start_col):
    if row_idx is None: return "???"
    return "".join([clean_cell(df.iloc[row_idx, c]) if clean_cell(df.iloc[row_idx, c]) != "" else "?" for c in range(start_col, start_col + 3)])

def get_three_digit_raw_list(df, row_idx, start_col):
    if row_idx is None: return ["?", "?", "?"]
    values = []
    for c in range(start_col, start_col + 3):
        v = clean_cell(df.iloc[row_idx, c])
        if v == "": v = "?"
        values.append(v)
    return values

def get_tooth_start_columns(df, tooth_row_idx):
    return {int(clean_cell(df.iloc[tooth_row_idx, col])): col for col in range(df.shape[1]) if is_valid_tooth(df.iloc[tooth_row_idx, col])}

def get_line_labels(tooth):
    q = tooth // 10
    if q == 1: return {"left_label_1": "DB", "right_label_1": "MB", "left_label_2": "DP", "right_label_2": "MP"}
    if q == 2: return {"left_label_1": "MB", "right_label_1": "DB", "left_label_2": "MP", "right_label_2": "DP"}
    if q == 3: return {"left_label_1": "ML", "right_label_1": "DL", "left_label_2": "MB", "right_label_2": "DB"}
    if q == 4: return {"left_label_1": "DL", "right_label_1": "ML", "left_label_2": "DB", "right_label_2": "MB"}
    raise ValueError(f"Invalid tooth: {tooth}")

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
    """
    智慧型列號收集器：支援單期 (Initial) 與雙期 (I vs R) 模式。
    target_stage: "I" (Initial) 或 "R" (Re-evaluation)
    """
    is_comp = is_comparison_file(df)
    pd_rows, gm_rows, mob_rows, km_rows = [], [], [], []

    for i in range(len(df)):
        row_prefix = " ".join([clean_cell(df.iloc[i, c]).upper() for c in range(min(3, df.shape[1]))])
        stage_tag = clean_cell(df.iloc[i, 1]).upper() if df.shape[1] > 1 else ""

        if is_comp:
            if "PD" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "PD" in row_prefix)):
                pd_rows.append(i if target_stage == "I" else i + 1)
            elif "GM" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "GM" in row_prefix)):
                gm_rows.append(i if target_stage == "I" else i + 1)
            elif "MOBILITY" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "MOBILITY" in row_prefix)):
                mob_rows.append(i if target_stage == "I" else i + 1)
            elif "KM" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "KM" in row_prefix)):
                km_rows.append(i if target_stage == "I" else i + 1)
        else:
            if "PD" in row_prefix: pd_rows.append(i)
            elif "GM" in row_prefix: gm_rows.append(i)
            elif "MOBILITY" in row_prefix: mob_rows.append(i)
            elif "KM" in row_prefix: km_rows.append(i)

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

def extract_implant_teeth(df_raw) -> set:
    """從 Charting CSV 中自動搜尋並解析所有被勾選為 Implant (植體) 的牙位"""
    tooth_rows = []
    implant_rows = []
    
    for idx, val in enumerate(df_raw[1]):
        if isinstance(val, str):
            if val.strip() == 'Tooth':
                tooth_rows.append(idx)
            elif val.strip() == 'Implant':
                implant_rows.append(idx)
                
    implant_teeth = set()
    
    if len(tooth_rows) >= 2 and len(implant_rows) >= 2:
        up_tooth_row = tooth_rows[0]
        up_implant_row = implant_rows[0]
        for col_idx in range(len(df_raw.columns) - 1):
            t_val = df_raw.iloc[up_tooth_row, col_idx]
            if isinstance(t_val, str) and t_val.isdigit():
                t = int(t_val)
                imp_val = str(df_raw.iloc[up_implant_row, col_idx + 1]).strip().upper()
                if imp_val == 'TRUE':
                    implant_teeth.add(t)
                    
        lo_tooth_row = tooth_rows[-1]
        lo_implant_row = implant_rows[-1]
        for col_idx in range(len(df_raw.columns) - 1):
            t_val = df_raw.iloc[lo_tooth_row, col_idx]
            if isinstance(t_val, str) and t_val.isdigit():
                t = int(t_val)
                imp_val = str(df_raw.iloc[lo_implant_row, col_idx + 1]).strip().upper()
                if imp_val == 'TRUE':
                    implant_teeth.add(t)
                    
    return implant_teeth

# ==============================================================================
# 🚀 對外進入點接口：完美串接 app.py 與 pptx_engine.py
# ==============================================================================

def parse_periodontal_csv(file_stream) -> Tuple[Any, Set[int], bool, Dict[str, Any]]:
    """
    對外主進入點：讀取 CSV/Excel 並調用上方經典台大解析邏輯。
    """
    try:
        file_stream.seek(0)
        df = pd.read_csv(file_stream, header=None)
    except Exception:
        file_stream.seek(0)
        df = pd.read_excel(file_stream, header=None)

    # 1. 判斷是否為雙期對比檔
    is_comp = is_comparison_file(df)

    # 2. 擷取病患 metadata
    patient_info = {}
    for r in range(min(10, len(df))):
        row = df.iloc[r].dropna().tolist()
        for idx, val in enumerate(row):
            val_str = str(val).strip()
            if "Patient's Name" in val_str and idx + 2 < len(row):
                patient_info["name"] = str(row[idx + 2]).strip()
            elif "Case Report No." in val_str and idx + 2 < len(row):
                patient_info["case_no"] = str(row[idx + 2]).strip()

    # 3. 呼叫專科邏輯計算 Missing Teeth
    tooth_rows = find_tooth_rows(df)
    missing_rows = find_missing_rows(df)
    missing_teeth = get_missing_teeth_set(df, tooth_rows, missing_rows)

    # 回傳原生 df，讓 pptx_engine 直接調用你的 get_three_digit_raw_list、collect_absolute_row_indices 等極致演算法
    return df, missing_teeth, is_comp, patient_info
