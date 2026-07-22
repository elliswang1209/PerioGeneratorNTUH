# core_parser.py
"""
核心解析引擎：自動相容單期 (Initial.csv) 與雙期 (Initial & Re-evaluation.csv) 牙周病歷格式。
100% 完整保留台大牙周專科病歷格式生成演算法，確保數據鏡射與過濾精準。
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
        # 尋找含有 Tooth 關鍵字或者包含超過 3 個有效牙位的列
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

        # 如果是雙期檔，根據 target_stage 篩選對應列
        if is_comp:
            # 第一個標題行通常帶有項目名稱，標記欄為 I
            # 第二行無項目名稱，標記欄為 R
            if "PD" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "PD" in row_prefix)):
                pd_rows.append(i if target_stage == "I" else i + 1)
            elif "GM" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "GM" in row_prefix)):
                gm_rows.append(i if target_stage == "I" else i + 1)
            elif "MOBILITY" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "MOBILITY" in row_prefix)):
                mob_rows.append(i if target_stage == "I" else i + 1)
            elif "KM" in row_prefix and (stage_tag == target_stage or (target_stage == "I" and "KM" in row_prefix)):
                km_rows.append(i if target_stage == "I" else i + 1)
        else:
            # 單期檔正常搜尋
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

def generate_present_dentition(df, tooth_rows, missing_teeth):
    if len(tooth_rows) < 2: return " 1. Present dentition\n    (Error: No tooth rows)"
    present = {t for t in get_tooth_start_columns(df, tooth_rows[0]).keys() if t not in missing_teeth} | {t for t in get_tooth_start_columns(df, tooth_rows[-1]).keys() if t not in missing_teeth}
    q1 = "".join([str(t % 10) if t in present else " " for t in [18, 17, 16, 15, 14, 13, 12, 11]])
    q2 = "".join([str(t % 10) if t in present else " " for t in [21, 22, 23, 24, 25, 26, 27, 28]])
    q4 = "".join([str(t % 10) if t in present else " " for t in [48, 47, 46, 45, 44, 43, 42, 41]])
    q3 = "".join([str(t % 10) if t in present else " " for t in [31, 32, 33, 34, 35, 36, 37, 38]])
    return f"LF:\n 1. Present dentition\n    {q1} | {q2}\n    {'-' * (len(q1)+len(q2)+3)}\n    {q4} | {q3}"

def parse_periodontal_pd(df, missing_teeth, target_stage="I"):
    tooth_rows = find_tooth_rows(df)
    anatomy = collect_absolute_row_indices(df, target_stage=target_stage)
    
    records = {}
    for tooth, col in get_tooth_start_columns(df, tooth_rows[0]).items():
        if tooth not in missing_teeth and tooth // 10 in [1, 2]:
            records[tooth] = {
                "tooth": tooth, 
                "col": col, 
                "buccal_or_facial": get_three_digit_pd(df, anatomy["up_b_pd"], col), 
                "palatal_or_lingual": get_three_digit_pd(df, anatomy["up_p_pd"], col)
            }
    for tooth, col in get_tooth_start_columns(df, tooth_rows[-1]).items():
        if tooth not in missing_teeth and tooth // 10 in [3, 4]:
            records[tooth] = {
                "tooth": tooth, 
                "col": col, 
                "buccal_or_facial": get_three_digit_pd(df, anatomy["lo_b_pd"], col), 
                "palatal_or_lingual": get_three_digit_pd(df, anatomy["lo_l_pd"], col)
            }
    return records

def parse_mobility(df, missing_teeth, target_stage="I"):
    tooth_rows = find_tooth_rows(df)
    anatomy = collect_absolute_row_indices(df, target_stage=target_stage)
    mobility = {}
    for t, c in get_tooth_start_columns(df, tooth_rows[0]).items():
        if t not in missing_teeth and anatomy["up_mob"] is not None: 
            mobility[t] = clean_cell(df.iloc[anatomy["up_mob"], c])
    for t, c in get_tooth_start_columns(df, tooth_rows[-1]).items():
        if t not in missing_teeth and anatomy["lo_mob"] is not None: 
            mobility[t] = clean_cell(df.iloc[anatomy["lo_mob"], c])
    return mobility

def parse_furcation(df, missing_teeth):
    tooth_rows = find_tooth_rows(df)
    f_rows = find_furcation_rows(df)
    if len(f_rows) < 2: return {}
    furcation = {}
    upper_prefered_order = ["M", "B", "D"]
    for tooth, col in get_tooth_start_columns(df, tooth_rows[0]).items():
        if tooth in missing_teeth: continue
        raw = {clean_cell(df.iloc[f_rows[0]["label_row"], col + o]).upper(): clean_cell(df.iloc[f_rows[0]["value_row"], col + o]) for o in range(3)}
        text = "".join([f"{l}{raw[l]}" for l in upper_prefered_order + ["L"] if l in raw and raw[l] in ["1", "2", "3"]])
        if text: furcation[tooth] = text
    lower_prefered_order = ["B", "L"]
    for tooth, col in get_tooth_start_columns(df, tooth_rows[-1]).items():
        if tooth in missing_teeth: continue
        raw = {clean_cell(df.iloc[f_rows[-1]["label_row"], col + o]).upper(): clean_cell(df.iloc[f_rows[-1]["value_row"], col + o]) for o in range(3)}
        text = "".join([f"{l}{raw[l]}" for l in lower_prefered_order + ["M", "D"] if l in raw and raw[l] in ["1", "2", "3"]])
        if text: furcation[tooth] = text
    return furcation
