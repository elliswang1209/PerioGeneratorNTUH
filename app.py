import streamlit as st
import pandas as pd
import io
from contextlib import redirect_stdout

# ============================================================
# 1. 牙周核心解析演算法 (保持完全不變，確保數據鏡射與過濾精準)
# ============================================================
def clean_cell(x): return str(x).strip()

def is_valid_tooth(x):
    x = clean_cell(x)
    if not x.isdigit(): return False
    tooth = int(x)
    return (tooth // 10) in [1, 2, 3, 4] and 1 <= (tooth % 10) <= 8

def find_tooth_rows(df):
    return [i for i in range(len(df)) if "Tooth" in [clean_cell(x) for x in df.iloc[i].tolist()]]

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
    return "".join([clean_cell(df.iloc[row_idx, c]) if clean_cell(df.iloc[row_idx, c]) != "" else "?" for c in range(start_col, start_col + 3)])

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

def generate_present_dentition(df, tooth_rows, missing_teeth):
    if len(tooth_rows) < 2: return " 1. Present dentition\n    (Error: No tooth rows)"
    present = {t for t in get_tooth_start_columns(df, tooth_rows[0]).keys() if t not in missing_teeth} | {t for t in get_tooth_start_columns(df, tooth_rows[-1]).keys() if t not in missing_teeth}
    q1 = "".join([str(t % 10) if t in present else " " for t in [18, 17, 16, 15, 14, 13, 12, 11]])
    q2 = "".join([str(t % 10) if t in present else " " for t in [21, 22, 23, 24, 25, 26, 27, 28]])
    q4 = "".join([str(t % 10) if t in present else " " for t in [48, 47, 46, 45, 44, 43, 42, 41]])
    q3 = "".join([str(t % 10) if t in present else " " for t in [31, 32, 33, 34, 35, 36, 37, 38]])
    return f"LF:\n 1. Present dentition\n    {q1} | {q2}\n    {'-' * (len(q1)+len(q2)+3)}\n    {q4} | {q3}"

def parse_periodontal_pd(df, missing_teeth):
    tooth_rows = find_tooth_rows(df)
    pd_rows = find_pd_rows(df)
    if len(tooth_rows) < 2 or len(pd_rows) != 4: raise ValueError("Invalid PD structure")
    records = {}
    for tooth, col in get_tooth_start_columns(df, tooth_rows[0]).items():
        if tooth not in missing_teeth and tooth // 10 in [1, 2]:
            records[tooth] = {"tooth": tooth, "col": col, "buccal_or_facial": get_three_digit_pd(df, pd_rows[0], col), "palatal_or_lingual": get_three_digit_pd(df, pd_rows[1], col)}
    for tooth, col in get_tooth_start_columns(df, tooth_rows[-1]).items():
        if tooth not in missing_teeth and tooth // 10 in [3, 4]:
            records[tooth] = {"tooth": tooth, "col": col, "buccal_or_facial": get_three_digit_pd(df, pd_rows[3], col), "palatal_or_lingual": get_three_digit_pd(df, pd_rows[2], col)}
    return records

def parse_mobility(df, missing_teeth):
    tooth_rows = find_tooth_rows(df)
    m_rows = find_mobility_rows(df)
    if len(m_rows) < 2: return {}
    mobility = {}
    for t, c in get_tooth_start_columns(df, tooth_rows[0]).items():
        if t not in missing_teeth: mobility[t] = clean_cell(df.iloc[m_rows[0], c])
    for t, c in get_tooth_start_columns(df, tooth_rows[-1]).items():
        if t not in missing_teeth: mobility[t] = clean_cell(df.iloc[m_rows[-1], c])
    return mobility

def parse_furcation(df, missing_teeth):
    tooth_rows = find_tooth_rows(df)
    f_rows = find_furcation_rows(df)
    if len(f_rows) < 2: return {}
    furcation = {}
    
    # Upper
    upper_prefered_order = ["M", "B", "D"]
    for tooth, col in get_tooth_start_columns(df, tooth_rows[0]).items():
        if tooth in missing_teeth: continue
        raw = {clean_cell(df.iloc[f_rows[0]["label_row"], col + o]).upper(): clean_cell(df.iloc[f_rows[0]["value_row"], col + o]) for o in range(3)}
        text = "".join([f"{l}{raw[l]}" for l in upper_prefered_order + ["L"] if l in raw and raw[l] in ["1", "2", "3"]])
        if text: furcation[tooth] = text
        
    # Lower
    lower_prefered_order = ["B", "L"]
    for tooth, col in get_tooth_start_columns(df, tooth_rows[-1]).items():
        if tooth in missing_teeth: continue
        raw = {clean_cell(df.iloc[f_rows[-1]["label_row"], col + o]).upper(): clean_cell(df.iloc[f_rows[-1]["value_row"], col + o]) for o in range(3)}
        text = "".join([f"{l}{raw[l]}" for l in lower_prefered_order + ["M", "D"] if l in raw and raw[l] in ["1", "2", "3"]])
        if text: furcation[tooth] = text
    return furcation

def format_single_tooth(record):
    tooth = record["tooth"]
    labels = get_line_labels(tooth)
    q = tooth // 10
    b_val, i_val = record["buccal_or_facial"], record["palatal_or_lingual"]
    l1, l2 = (b_val, i_val) if q in [1, 2] else (i_val, b_val)
    return [f"tooth {tooth}", f"   {labels['left_label_1']} {l1} {labels['right_label_1']}", f"   {labels['left_label_2']} {l2} {labels['right_label_2']}"]

def print_teeth_side_by_side(records, tooth_list, gap=10):
    blocks = [format_single_tooth(records[t]) for t in tooth_list if t in records]
    if not blocks: return
    widths = [max(len(line) for line in block) for block in blocks]
    for row_idx in range(3):
        print((" " * gap).join([block[row_idx].ljust(width) for block, width in zip(blocks, widths)]))
    print()

def get_flagged_teeth_string(records, ordered_groups, gap=6):
    f = io.StringIO()
    flagged = {t for t, r in records.items() if any(c.isdigit() and int(c) >= 5 for c in f"{r['buccal_or_facial']}{r['palatal_or_lingual']}") or "?" in f"{r['buccal_or_facial']}{r['palatal_or_lingual']}"}
    flagged_list_str = " ".join([str(t) for g in ordered_groups for t in g if t in flagged])
    with redirect_stdout(f):
        print(f" 3. Probing depth >=5mm: tooth {flagged_list_str}\n")
        for g in ordered_groups:
            to_print = [t for t in g if t in records and t in flagged]
            if to_print: print_teeth_side_by_side(records, to_print, gap=gap)
    return f.getvalue()

def print_mobility_summary(df, missing_teeth):
    mobility = parse_mobility(df, missing_teeth)
    g1, g2, g3 = [], [], []
    for t, g in mobility.items():
        if g == "1": g1.append(str(t))
        elif g == "2": g2.append(str(t))
        elif g == "3": g3.append(str(t))
    print(" 4. Mobility:")
    print("-Gr. I: " + ("nil" if not g1 else "tooth " + " ".join(g1)))
    print("-Gr. II: " + ("nil" if not g2 else "tooth " + " ".join(g2)))
    print("-Gr. III: " + ("nil" if not g3 else "tooth " + " ".join(g3)))

def print_furcation_summary(df, missing_teeth):
    furcation = parse_furcation(df, missing_teeth)
    print("\n 5. Furcation:")
    up, lo = [], []
    for t in sorted(furcation.keys()):
        (up if (t // 10 in [1, 2]) else lo).append(f"{t}({furcation[t]})")
    print("-Upper: " + ("nil" if not up else "tooth " + " ".join(up)))
    print("-Lower: " + ("nil" if not lo else "tooth " + " ".join(lo)))


# ============================================================
# 2. Streamlit 前端介面設計 (寫死唯讀版 + 清除功能)
# ============================================================

st.set_page_config(page_title="牙周病歷自動生成器", page_icon="🦷", layout="centered")

st.title("🦷 Perio Generator for NTUH")
st.markdown("可以把 Charting Data 轉換成病歷的文字格式。")

# 🌟 為上傳元件加上 key，以便後續一鍵清除
uploaded_file = st.file_uploader(
    "📂 請上傳轉換成為「CSV檔案格式」的 Charting Excel", 
    type=["csv"],
    key="perio_csv_uploader"
)

if uploaded_file is not None:
    try:
        df_raw = pd.read_csv(uploaded_file, header=None, dtype=str).fillna("")
        
        tooth_rows_idx = find_tooth_rows(df_raw)
        missing_rows_idx = find_missing_rows(df_raw)
        missing_teeth = get_missing_teeth_set(df_raw, tooth_rows_idx, missing_rows_idx)
        records = parse_periodontal_pd(df_raw, missing_teeth)
        
        ordered_groups = [
            [18, 17, 16, 15, 14, 13, 12, 11],
            [21, 22, 23, 24, 25, 26, 27, 28],
            [38, 37, 36, 35, 34, 33, 32, 31],
            [41, 42, 43, 44, 45, 46, 47, 48],
        ]
        
        output_buffer = io.StringIO()
        with redirect_stdout(output_buffer):
            print(generate_present_dentition(df_raw, tooth_rows_idx, missing_teeth))
            print("")
            print(" 2. Full mouth general calculus and plaque deposition")
            print("")
            
            pd_report_str = get_flagged_teeth_string(records, ordered_groups, gap=6)
            print(pd_report_str.rstrip())
            print("")
            
            try: print_mobility_summary(df_raw, missing_teeth)
            except: print(" 4. Mobility: nil")
            
            try: print_furcation_summary(df_raw, missing_teeth)
            except: print("\n 5. Furcation: nil")
            
        final_note = output_buffer.getvalue()
        
        st.success("🎉 病歷文字成功生成！")
        st.subheader("📋 臨床 Progress Note 文字")
        
        # 🌟 換回 st.code：不可編輯、格式固定、右上角自帶一鍵複製按鈕
        st.code(final_note, language="text")
        
        st.write("") # 增加一些垂直間距
        st.write("---")
        
        # 🌟 末端按鍵：生成下一位（清除 Panel 數據與上傳檔案）
        if st.button("🔄 生成下一位（清除目前資料）", type="primary", use_container_width=True):
            # 清除暫存與快取
            st.cache_data.clear()
            # 利用 Streamlit 的 rerun 機制重整網頁，完全回復初始上傳狀態
            st.rerun()
            
    except Exception as e:
        st.error(f"❌ 檔案解析失敗。請確認該 CSV 的結構是否正確。詳細錯誤：{e}")
