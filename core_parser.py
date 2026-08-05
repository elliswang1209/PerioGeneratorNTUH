"""
核心解析引擎：100% 完整保留台大牙周專科病歷格式生成演算法，確保數據鏡射與過濾精準。
支援 Initial 與 Initial & Re-evaluation 雙期分側定位，修正 CAL 被 MOBILITY SCALE 誤抓的問題。
"""
import pandas as pd
import io
import re
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

def find_missing_teeth(df) -> Set[int]:
    """快捷缺牙掃描函式，確保 app.py 呼叫不拋錯"""
    tooth_rows = find_tooth_rows(df)
    missing_rows = find_missing_rows(df)
    return get_missing_teeth_set(df, tooth_rows, missing_rows)

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

# ============================================================
# 🚀 補全 app.py 介面相容函式 (Objective 病歷文字編織)
# ============================================================

def collect_absolute_row_indices(df: pd.DataFrame) -> dict:
    """相容性函式：提供 app.py 與 pptx_engine 解鎖特定資料行號"""
    return collect_comparison_row_indices(df)

def extract_implant_teeth(df: pd.DataFrame) -> set:
    """提取植牙牙號集合"""
    implant_teeth = set()
    for r in range(len(df)):
        row_str = " ".join([clean_cell(df.iloc[r, c]).upper() for c in range(df.shape[1])])
        if "IMPLANT" in row_str:
            tooth_rows = find_tooth_rows(df)
            if tooth_rows:
                start_cols = get_tooth_start_columns(df, tooth_rows[0])
                for t, c in start_cols.items():
                    val = clean_cell(df.iloc[r, c + 1]).upper() if c + 1 < df.shape[1] else ""
                    if val == "TRUE" or "IMPLANT" in val:
                        implant_teeth.add(t)
    return implant_teeth

def parse_periodontal_pd(df: pd.DataFrame, missing_teeth: set) -> dict:
    """解析 11~48 牙齒之 Probing Depth (PD) 口袋深度"""
    row_map = collect_comparison_row_indices(df)
    tooth_rows = find_tooth_rows(df)
    records = {}

    if not tooth_rows:
        return records

    up_cols = get_tooth_start_columns(df, tooth_rows[0])
    lo_cols = get_tooth_start_columns(df, tooth_rows[-1]) if len(tooth_rows) > 1 else {}

    for tooth in range(11, 49):
        if tooth in missing_teeth:
            continue

        pd_vals = []
        col_start = None

        if tooth // 10 in [1, 2] and tooth in up_cols:
            col_start = up_cols[tooth]
            r_buccal = row_map.get("up_b_pd_i")
            r_palatal = row_map.get("up_p_pd_i")
            raw_b = get_three_digit_raw_list(df, r_buccal, col_start)
            raw_p = get_three_digit_raw_list(df, r_palatal, col_start)
            for v in raw_b + raw_p:
                if v.isdigit(): pd_vals.append(int(v))

        elif tooth // 10 in [3, 4] and tooth in lo_cols:
            col_start = lo_cols[tooth]
            r_lingual = row_map.get("lo_l_pd_i")
            r_buccal = row_map.get("lo_b_pd_i")
            raw_l = get_three_digit_raw_list(df, r_lingual, col_start)
            raw_b = get_three_digit_raw_list(df, r_buccal, col_start)
            for v in raw_l + raw_b:
                if v.isdigit(): pd_vals.append(int(v))

        records[tooth] = {"pd_values": pd_vals}

    return records

def generate_present_dentition(df: pd.DataFrame, tooth_rows: list, missing_teeth: set) -> str:
    """編織 Present Dentition 病歷文字格式"""
    all_teeth = set(range(11, 19)) | set(range(21, 29)) | set(range(31, 39)) | set(range(41, 49))
    present = sorted(list(all_teeth - missing_teeth))
    missing = sorted(list(missing_teeth))

    present_str = ", ".join(map(str, present)) if present else "None"
    missing_str = ", ".join(map(str, missing)) if missing else "nil"

    return f" 1. Present dentition: {present_str}\n    Missing teeth: {missing_str}"

def get_flagged_teeth_string(records: dict, ordered_groups: list, gap: int = 6) -> str:
    """收集口袋深度 PD >= 4mm 的牙號報告"""
    flagged = []
    for tooth in sorted(records.keys()):
        pd_vals = records[tooth].get("pd_values", [])
        if any(v >= 4 for v in pd_vals):
            flagged.append(str(tooth))

    flagged_str = ", ".join(flagged) if flagged else "nil"
    return f" 3. Probing depth >= 4mm noted at teeth: {flagged_str}"

def print_mobility_summary(df: pd.DataFrame, missing_teeth: set):
    """印出 Mobility 搖動度摘要"""
    row_map = collect_comparison_row_indices(df)
    tooth_rows = find_tooth_rows(df)
    mob_teeth = []

    if tooth_rows:
        up_cols = get_tooth_start_columns(df, tooth_rows[0])
        r_up_mob = row_map.get("up_mob_i")
        if r_up_mob is not None:
            for t, c in up_cols.items():
                if t not in missing_teeth:
                    v = clean_cell(df.iloc[r_up_mob, c])
                    if v and v != "0" and v != "?":
                        mob_teeth.append(f"{t}(Degree {v})")

        if len(tooth_rows) > 1:
            lo_cols = get_tooth_start_columns(df, tooth_rows[-1])
            r_lo_mob = row_map.get("lo_mob_i")
            if r_lo_mob is not None:
                for t, c in lo_cols.items():
                    if t not in missing_teeth:
                        v = clean_cell(df.iloc[r_lo_mob, c])
                        if v and v != "0" and v != "?":
                            mob_teeth.append(f"{t}(Degree {v})")

    if mob_teeth:
        print(f" 4. Mobility: {', '.join(mob_teeth)}")
    else:
        print(" 4. Mobility: nil")

def print_furcation_summary(df: pd.DataFrame, missing_teeth: set):
    """印出 Furcation 根分叉摘要"""
    furc_rows = find_furcation_rows(df)
    tooth_rows = find_tooth_rows(df)
    furc_teeth = []

    if furc_rows and tooth_rows:
        up_cols = get_tooth_start_columns(df, tooth_rows[0])
        r_up_val = furc_rows[0]["value_row"] if len(furc_rows) > 0 else None

        if r_up_val is not None and r_up_val < len(df):
            for t, c in up_cols.items():
                if t not in missing_teeth:
                    raw_vals = get_three_digit_raw_list(df, r_up_val, c)
                    valid_f = [v for v in raw_vals if v in ['1', '2', '3', 'I', 'II', 'III']]
                    if valid_f:
                        furc_teeth.append(f"{t}(Class {valid_f[0]})")

    if furc_teeth:
        print(f"\n 5. Furcation: {', '.join(furc_teeth)}")
    else:
        print("\n 5. Furcation: nil")

# ============================================================
# 🚀 漂亮格式 Objective 病歷文字生成核心函式
# ============================================================

def generate_cross_dentition(missing_teeth: set) -> str:
    """生成 1. Present dentition 十字齒列圖形"""
    # 定義四個象限的所有牙齒個位數 (1~8)
    q1 = [t % 10 for t in range(18, 10, -1) if t not in missing_teeth]
    q2 = [t % 10 for t in range(21, 29) if t not in missing_teeth]
    q4 = [t % 10 for t in range(48, 40, -1) if t not in missing_teeth]
    q3 = [t % 10 for t in range(31, 39) if t not in missing_teeth]

    str_q1 = "".join(map(str, q1)).rjust(8)
    str_q2 = "".join(map(str, q2)).ljust(8)
    str_q4 = "".join(map(str, q4)).rjust(8)
    str_q3 = "".join(map(str, q3)).ljust(8)

    border = "-" * 19
    return (
        f" 1. Present dentition\n"
        f"    {str_q1} | {str_q2}\n"
        f"    {border}\n"
        f"    {str_q4} | {str_q3}"
    )

def format_detailed_pd_section(df: pd.DataFrame, missing_teeth: set, threshold: int = 5) -> str:
    """生成 3. Probing depth >= 5mm 詳細六點位點對齊排版 (每行最多4顆牙)"""
    records = parse_periodontal_pd(df, missing_teeth)
    row_map = collect_comparison_row_indices(df)
    tooth_rows = find_tooth_rows(df)

    if not tooth_rows:
        return f" 3. Probing depth >={threshold}mm: nil"

    up_cols = get_tooth_start_columns(df, tooth_rows[0])
    lo_cols = get_tooth_start_columns(df, tooth_rows[-1]) if len(tooth_rows) > 1 else {}

    flagged_teeth = []
    tooth_details = {}

    for tooth in sorted(records.keys()):
        pd_vals = records[tooth].get("pd_values", [])
        if any(v >= threshold for v in pd_vals if isinstance(v, int)):
            flagged_teeth.append(tooth)
            
            # 抓取 6 個位點數字，若缺漏則以 '?' 補足
            col_start = up_cols.get(tooth) if tooth // 10 in [1, 2] else lo_cols.get(tooth)
            
            if tooth // 10 in [1, 2]:
                r_b = row_map.get("up_b_pd_i")
                r_p = row_map.get("up_p_pd_i")
                b_vals = get_three_digit_raw_list(df, r_b, col_start) if col_start is not None else ["?", "?", "?"]
                p_vals = get_three_digit_raw_list(df, r_p, col_start) if col_start is not None else ["?", "?", "?"]
                
                # 上顎格式：MB, B, DB / MP, P, DP
                line1 = f"DB {b_vals[2]}{b_vals[1]}{b_vals[0]} MB" if tooth // 10 == 1 else f"MB {b_vals[0]}{b_vals[1]}{b_vals[2]} DB"
                line2 = f"DP {p_vals[2]}{p_vals[1]}{p_vals[0]} MP" if tooth // 10 == 1 else f"MP {p_vals[0]}{p_vals[1]}{p_vals[2]} DP"
            else:
                r_l = row_map.get("lo_l_pd_i")
                r_b = row_map.get("lo_b_pd_i")
                l_vals = get_three_digit_raw_list(df, r_l, col_start) if col_start is not None else ["?", "?", "?"]
                b_vals = get_three_digit_raw_list(df, r_b, col_start) if col_start is not None else ["?", "?", "?"]
                
                # 下顎格式
                line1 = f"DL {l_vals[2]}{l_vals[1]}{l_vals[0]} ML" if tooth // 10 == 4 else f"ML {l_vals[0]}{l_vals[1]}{l_vals[2]} DL"
                line2 = f"DB {b_vals[2]}{b_vals[1]}{b_vals[0]} MB" if tooth // 10 == 4 else f"MB {b_vals[0]}{b_vals[1]}{b_vals[2]} DB"

            tooth_details[tooth] = (line1, line2)

    if not flagged_teeth:
        return f" 3. Probing depth >={threshold}mm: nil"

    teeth_str = " ".join(map(str, flagged_teeth))
    out = [f" 3. Probing depth >={threshold}mm: tooth {teeth_str}\n"]

    # 每 4 顆牙分一組印出
    for i in range(0, len(flagged_teeth), 4):
        chunk = flagged_teeth[i:i+4]
        
        # 標題行 (例如: tooth 17          tooth 16 ...)
        header = "".join([f"tooth {t:<12}" for t in chunk])
        l1 = "".join([f"   {tooth_details[t][0]:<15}" for t in chunk])
        l2 = "".join([f"   {tooth_details[t][1]:<15}" for t in chunk])

        out.append(header)
        out.append(l1)
        out.append(l2)
        out.append("")  # 空行分隔組別

    return "\n".join(out)

def format_mobility_section(df: pd.DataFrame, missing_teeth: set) -> str:
    """生成 4. Mobility 分級報告"""
    row_map = collect_comparison_row_indices(df)
    tooth_rows = find_tooth_rows(df)

    gr1, gr2, gr3 = [], [], []

    if tooth_rows:
        up_cols = get_tooth_start_columns(df, tooth_rows[0])
        lo_cols = get_tooth_start_columns(df, tooth_rows[-1]) if len(tooth_rows) > 1 else {}
        
        for tooth, col in {**up_cols, **lo_cols}.items():
            if tooth in missing_teeth: continue
            r_mob = row_map.get("up_mob_i") if tooth // 10 in [1, 2] else row_map.get("lo_mob_i")
            if r_mob is not None:
                v = clean_cell(df.iloc[r_mob, col])
                if v in ["1", "I"]: gr1.append(str(tooth))
                elif v in ["2", "II"]: gr2.append(str(tooth))
                elif v in ["3", "III"]: gr3.append(str(tooth))

    str_g1 = " ".join(gr1) if gr1 else "nil"
    str_g2 = " ".join(gr2) if gr2 else "nil"
    str_g3 = " ".join(gr3) if gr3 else "nil"

    return f" 4. Mobility:\n-Gr. I: {str_g1}\n-Gr. II: {str_g2}\n-Gr. III: {str_g3}"

def format_furcation_section(df: pd.DataFrame, missing_teeth: set) -> str:
    """生成 5. Furcation 上下顎報告"""
    furc_rows = find_furcation_rows(df)
    tooth_rows = find_tooth_rows(df)

    upper_furc, lower_furc = [], []

    if furc_rows and tooth_rows:
        up_cols = get_tooth_start_columns(df, tooth_rows[0])
        lo_cols = get_tooth_start_columns(df, tooth_rows[-1]) if len(tooth_rows) > 1 else {}

        r_up = furc_rows[0]["value_row"] if len(furc_rows) > 0 else None
        r_lo = furc_rows[-1]["value_row"] if len(furc_rows) > 1 else None

        if r_up is not None and r_up < len(df):
            for t, c in up_cols.items():
                if t not in missing_teeth:
                    vals = get_three_digit_raw_list(df, r_up, c)
                    valid = [v for v in vals if v in ['1', '2', '3', 'I', 'II', 'III']]
                    if valid: upper_furc.append(f"{t}(Grade {valid[0]})")

        if r_lo is not None and r_lo < len(df):
            for t, c in lo_cols.items():
                if t not in missing_teeth:
                    vals = get_three_digit_raw_list(df, r_lo, c)
                    valid = [v for v in vals if v in ['1', '2', '3', 'I', 'II', 'III']]
                    if valid: lower_furc.append(f"{t}(Grade {valid[0]})")

    str_up = " ".join(upper_furc) if upper_furc else "nil"
    str_lo = " ".join(lower_furc) if lower_furc else "nil"

    return f" 5. Furcation:\n-Upper: {str_up}\n-Lower: {str_lo}"
