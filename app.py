import pandas as pd
import os
import sys


def clean_cell(x):
    return str(x).strip()


def is_valid_tooth(x):
    x = clean_cell(x)
    if not x.isdigit():
        return False

    tooth = int(x)
    quadrant = tooth // 10
    number = tooth % 10
    return quadrant in [1, 2, 3, 4] and 1 <= number <= 8


def find_tooth_rows(df):
    tooth_rows = []
    for i in range(len(df)):
        row_values = [clean_cell(x) for x in df.iloc[i].tolist()]
        if "Tooth" in row_values:
            tooth_rows.append(i)
    return tooth_rows


def find_missing_rows(df):
    missing_rows = []
    for i in range(len(df)):
        row_values = [clean_cell(x).upper() for x in df.iloc[i].tolist()]
        if "MISSING" in row_values:
            missing_rows.append(i)
    return missing_rows


def find_pd_rows(df):
    pd_rows = []
    for i in range(len(df)):
        row_values = [clean_cell(x) for x in df.iloc[i].tolist()]
        if any(x.upper().startswith("PD") for x in row_values):
            pd_rows.append(i)
    return pd_rows


def find_mobility_rows(df):
    mobility_rows = []
    for i in range(len(df)):
        row_values = [clean_cell(x) for x in df.iloc[i].tolist()]
        if any("Mobility" in x for x in row_values):
            mobility_rows.append(i)
    return mobility_rows


def find_furcation_rows(df):
    furcation_rows = []
    for i in range(len(df)):
        row_values = [clean_cell(x) for x in df.iloc[i].tolist()]
        row_text = " ".join(row_values).lower()
        if "furcation" in row_text:
            furcation_rows.append({
                "header_row": i,
                "label_row": i,
                "value_row": i + 1
            })
    return furcation_rows


def get_three_digit_pd(df, row_idx, start_col):
    values = []
    for c in range(start_col, start_col + 3):
        v = clean_cell(df.iloc[row_idx, c])
        if v == "":
            v = "?"
        values.append(v)
    return "".join(values)


def get_tooth_start_columns(df, tooth_row_idx):
    tooth_cols = {}
    for col in range(df.shape[1]):
        value = clean_cell(df.iloc[tooth_row_idx, col])
        if is_valid_tooth(value):
            tooth_cols[int(value)] = col
    return tooth_cols


def get_line_labels(tooth):
    """
    根據四象限臨床解剖位置，回傳最精準的左右標籤配置（實現鏡射邏輯）
    """
    q = tooth // 10

    # Q1 (右上): 左=Distal(DB/DP), 右=Mesial(MB/MP) -> [原始設定]
    if q == 1:
        return {
            "left_label_1": "DB", "right_label_1": "MB",
            "left_label_2": "DP", "right_label_2": "MP"
        }

    # Q2 (左上): 左=Mesial(MB/MP), 右=Distal(DB/DP) -> [左右鏡射]
    if q == 2:
        return {
            "left_label_1": "MB", "right_label_1": "DB",
            "left_label_2": "MP", "right_label_2": "DP"
        }

    # Q3 (左下): 上排(1)=Lingual(ML/DL), 下排(2)=Buccal(MB/DB)，且左=Mesial, 右=Distal -> [上下翻轉 + 左右鏡射]
    if q == 3:
        return {
            "left_label_1": "ML", "right_label_1": "DL",
            "left_label_2": "MB", "right_label_2": "DB"
        }

    # Q4 (右下): 上排(1)=Lingual(DL/ML), 下排(2)=Buccal(DB/MB)，且左=Distal, 右=Mesial -> [上下翻轉]
    if q == 4:
        return {
            "left_label_1": "DL", "right_label_1": "ML",
            "left_label_2": "DB", "right_label_2": "MB"
        }

    raise ValueError(f"Invalid tooth number: {tooth}")

def get_missing_teeth_set(df, tooth_rows, missing_rows):
    """
    從 CSV 中分析並回傳所有被標記為 Missing (TRUE) 的牙齒集合
    """
    missing_teeth = set()
    if len(tooth_rows) < 2:
        return missing_teeth

    upper_tooth_row = tooth_rows[0]
    lower_tooth_row = tooth_rows[-1]
    upper_missing_row = missing_rows[0] if len(missing_rows) > 0 else None
    lower_missing_row = missing_rows[-1] if len(missing_rows) > 1 else (missing_rows[0] if len(missing_rows) == 1 else None)

    upper_cols = get_tooth_start_columns(df, upper_tooth_row)
    lower_cols = get_tooth_start_columns(df, lower_tooth_row)

    # 檢查上顎
    for tooth, col in upper_cols.items():
        if tooth // 10 in [1, 2] and upper_missing_row is not None and col + 1 < df.shape[1]:
            if clean_cell(df.iloc[upper_missing_row, col + 1]).upper() == "TRUE":
                missing_teeth.add(tooth)

    # 檢查下顎
    for tooth, col in lower_cols.items():
        if tooth // 10 in [3, 4] and lower_missing_row is not None and col + 1 < df.shape[1]:
            if clean_cell(df.iloc[lower_missing_row, col + 1]).upper() == "TRUE":
                missing_teeth.add(tooth)

    return missing_teeth


def generate_present_dentition(df, tooth_rows, missing_teeth):
    """
    利用已知的 missing_teeth 集合，生成正確挖空的十字牙位圖。
    """
    if len(tooth_rows) < 2:
        return " 1. Present dentition\n    (Error: Unable to find tooth rows)"

    upper_cols = get_tooth_start_columns(df, tooth_rows[0])
    lower_cols = get_tooth_start_columns(df, tooth_rows[-1])

    # 只要出現在欄位中，且不在 missing_teeth 內，就是 Present
    present_teeth = set()
    for t in upper_cols.keys():
        if t not in missing_teeth:
            present_teeth.add(t)
    for t in lower_cols.keys():
        if t not in missing_teeth:
            present_teeth.add(t)

    q1_seq = [18, 17, 16, 15, 14, 13, 12, 11]
    q2_seq = [21, 22, 23, 24, 25, 26, 27, 28]
    q4_seq = [48, 47, 46, 45, 44, 43, 42, 41]
    q3_seq = [31, 32, 33, 34, 35, 36, 37, 38]

    q1_str = "".join([str(t % 10) if t in present_teeth else " " for t in q1_seq])
    q2_str = "".join([str(t % 10) if t in present_teeth else " " for t in q2_seq])
    q4_str = "".join([str(t % 10) if t in present_teeth else " " for t in q4_seq])
    q3_str = "".join([str(t % 10) if t in present_teeth else " " for t in q3_seq])

    line_len = len(q1_str) + len(q2_str) + 3

    output = []
    output.append(f" 1. Present dentition")
    output.append(f"    {q1_str} | {q2_str}")
    output.append(f"    " + "-" * line_len)
    output.append(f"    {q4_str} | {q3_str}")
    return "\n".join(output)


# ============================================================
# Main parsers with Missing Filters
# ============================================================

def parse_periodontal_pd(csv_path, missing_teeth):
    df = pd.read_csv(csv_path, header=None, dtype=str).fillna("")

    tooth_rows = find_tooth_rows(df)
    pd_rows = find_pd_rows(df)

    if len(tooth_rows) < 2 or len(pd_rows) != 4:
        raise ValueError("Invalid CSV structure for PD rows")

    upper_cols = get_tooth_start_columns(df, tooth_rows[0])
    lower_cols = get_tooth_start_columns(df, tooth_rows[-1])

    records = {}

    # 上顎
    for tooth, col in upper_cols.items():
        if tooth in missing_teeth or tooth // 10 not in [1, 2]:
            continue  # 🌟 缺失牙直接跳過，不列入 PD 統計
        buccal = get_three_digit_pd(df, pd_rows[0], col)
        inner = get_three_digit_pd(df, pd_rows[1], col)
        records[tooth] = {
            "tooth": tooth,
            "col": col,
            "buccal_or_facial": buccal,
            "palatal_or_lingual": inner,
        }

    # 下顎
    for tooth, col in lower_cols.items():
        if tooth in missing_teeth or tooth // 10 not in [3, 4]:
            continue  # 🌟 缺失牙直接跳過，不列入 PD 統計
        buccal = get_three_digit_pd(df, pd_rows[3], col)
        inner = get_three_digit_pd(df, pd_rows[2], col)
        records[tooth] = {
            "tooth": tooth,
            "col": col,
            "buccal_or_facial": buccal,
            "palatal_or_lingual": inner,
        }

    return records


def parse_mobility(csv_path, missing_teeth):
    df = pd.read_csv(csv_path, header=None, dtype=str).fillna("")
    tooth_rows = find_tooth_rows(df)
    mobility_rows = find_mobility_rows(df)

    if len(mobility_rows) < 2:
        return {}

    upper_cols = get_tooth_start_columns(df, tooth_rows[0])
    lower_cols = get_tooth_start_columns(df, tooth_rows[-1])

    mobility = {}
    for tooth, col in upper_cols.items():
        if tooth in missing_teeth:
            continue  # 🌟 缺失牙跳過
        mobility[tooth] = clean_cell(df.iloc[mobility_rows[0], col])
        
    for tooth, col in lower_cols.items():
        if tooth in missing_teeth:
            continue  # 🌟 缺失牙跳過
        mobility[tooth] = clean_cell(df.iloc[mobility_rows[-1], col])
        
    return mobility


def parse_furcation(csv_path, missing_teeth):
    df = pd.read_csv(csv_path, header=None, dtype=str).fillna("")
    tooth_rows = find_tooth_rows(df)
    furcation_rows = find_furcation_rows(df)

    if len(furcation_rows) < 2:
        return {}

    upper_cols = get_tooth_start_columns(df, tooth_rows[0])
    lower_cols = get_tooth_start_columns(df, tooth_rows[-1])

    upper_furcation = furcation_rows[0]
    lower_furcation = furcation_rows[-1]

    furcation = {}

    # ===================================
    # Upper teeth (上顎：依照 M -> B -> D 順序排列)
    # ===================================
    upper_prefered_order = ["M", "B", "D"]
    for tooth, col in upper_cols.items():
        if tooth in missing_teeth:
            continue
        
        # 先抓取這顆牙齒所有被填入的 furcation 數值
        raw_data = {}
        for offset in range(3):
            label = clean_cell(df.iloc[upper_furcation["label_row"], col + offset]).upper()
            value = clean_cell(df.iloc[upper_furcation["value_row"], col + offset])
            if label in ["D", "M", "B", "L"] and value in ["1", "2", "3"]:
                raw_data[label] = value
        
        # 按照 M -> B -> D -> (L防呆) 的順序重新組裝字串
        text = ""
        for label in upper_prefered_order + ["L"]:
            if label in raw_data:
                text += f"{label}{raw_data[label]}"
                
        if text:
            furcation[tooth] = text

    # ===================================
    # Lower teeth (下顎：依照 B -> L 順序排列)
    # ===================================
    lower_prefered_order = ["B", "L"]
    for tooth, col in lower_cols.items():
        if tooth in missing_teeth:
            continue
            
        raw_data = {}
        for offset in range(3):
            label = clean_cell(df.iloc[lower_furcation["label_row"], col + offset]).upper()
            value = clean_cell(df.iloc[lower_furcation["value_row"], col + offset])
            if label in ["D", "M", "B", "L"] and value in ["1", "2", "3"]:
                raw_data[label] = value
        
        # 按照 B -> L -> (M/D防呆) 的順序重新組裝字串
        text = ""
        for label in lower_prefered_order + ["M", "D"]:
            if label in raw_data:
                text += f"{label}{raw_data[label]}"
                
        if text:
            furcation[tooth] = text

    return furcation

# ============================================================
# Output Helper Functions
# ============================================================

def format_single_tooth(record):
    """
    依據四象限鏡射邏輯，將 3 位數的 PD 數值進行對應的左右鏡射與上下調換。
    """
    tooth = record["tooth"]
    labels = get_line_labels(tooth)
    q = tooth // 10

    buccal_val = record["buccal_or_facial"]    # 原始 CSV 順序通常是固定方向
    inner_val = record["palatal_or_lingual"]   # 原始 CSV 順序通常是固定方向

    if q in [1, 2]:
        # 上顎：第一行印 Facial/Buccal，第二行印 Palatal
        line1_num = buccal_val
        line2_num = inner_val
    else:
        # 下顎：第一行印 Lingual，第二行印 Buccal
        line1_num = inner_val
        line2_num = buccal_val

    # 組裝終端機呈現畫面
    line0 = f"tooth {tooth}"
    line1 = f"   {labels['left_label_1']} {line1_num} {labels['right_label_1']}"
    line2 = f"   {labels['left_label_2']} {line2_num} {labels['right_label_2']}"

    return [line0, line1, line2]


def print_teeth_side_by_side(records, tooth_list, gap=10):
    blocks = []
    for tooth in tooth_list:
        if tooth not in records:
            continue
        blocks.append(format_single_tooth(records[tooth]))

    if not blocks:
        return

    widths = [max(len(line) for line in block) for block in blocks]
    for row_idx in range(3):
        row_parts = []
        for block, width in zip(blocks, widths):
            row_parts.append(block[row_idx].ljust(width))
        print((" " * gap).join(row_parts))
    print()


def get_flagged_teeth_string(records, ordered_groups, gap=6):
    import io
    from contextlib import redirect_stdout

    f = io.StringIO()
    flagged = set()
    for tooth, rec in records.items():
        six = f"{rec.get('buccal_or_facial','')}" + f"{rec.get('palatal_or_lingual','')}"
        if any(ch.isdigit() and int(ch) >= 5 for ch in six) or "?" in six:
            flagged.add(tooth)

    flagged_list_str = " ".join([str(t) for grp in ordered_groups for t in grp if t in flagged])
    
    with redirect_stdout(f):
        print(f"3. Probing depth >=5mm: tooth {flagged_list_str}\n")
        
        for grp in ordered_groups:
            to_print = [t for t in grp if t in records and t in flagged]
            if to_print:
                print_teeth_side_by_side(records, to_print, gap=gap)
                
    return f.getvalue()


def print_mobility_summary(csv_path, missing_teeth):
    mobility = parse_mobility(csv_path, missing_teeth)
    grade1, grade2, grade3 = [], [], []

    for tooth, grade in mobility.items():
        grade = grade.strip()
        if grade == "1":
            grade1.append(str(tooth))
        elif grade == "2":
            grade2.append(str(tooth))
        elif grade == "3":
            grade3.append(str(tooth))

    print("4. Mobility:")
    print("-Gr. I: " + ("nil" if not grade1 else "tooth " + " ".join(grade1)))
    print("-Gr. II: " + ("nil" if not grade2 else "tooth " + " ".join(grade2)))
    print("-Gr. III: " + ("nil" if not grade3 else "tooth " + " ".join(grade3)))


def print_furcation_summary(csv_path, missing_teeth):
    furcation = parse_furcation(csv_path, missing_teeth)
    
    print("\n5. Furcation:")

    upper_list = []
    lower_list = []
    
    for tooth in sorted(furcation.keys()):
        quadrant = tooth // 10
        item_str = f"{tooth}({furcation[tooth]})"
        
        if quadrant in [1, 2]:
            upper_list.append(item_str)
        elif quadrant in [3, 4]:
            lower_list.append(item_str)
            
    if upper_list:
        print(f"-Upper: tooth " + " ".join(upper_list))
    else:
        print("-Upper: nil")
        
    if lower_list:
        print(f"-Lower: tooth " + " ".join(lower_list))
    else:
        print("-Lower: nil")


if __name__ == "__main__":

    csv_path = "/Users/wangguanzhong/Desktop/大六 Intern/2_PR/余承豐_6755492_Q2 - Initial(輸入).csv"

    if not os.path.exists(csv_path):
        print(f"❌ 錯誤：找不到 CSV 檔案：\n{csv_path}")
        sys.exit(1)

    df_raw = pd.read_csv(csv_path, header=None, dtype=str).fillna("")
    tooth_rows_idx = find_tooth_rows(df_raw)
    missing_rows_idx = find_missing_rows(df_raw)
    missing_teeth = get_missing_teeth_set(df_raw, tooth_rows_idx, missing_rows_idx)

    try:
        records = parse_periodontal_pd(csv_path, missing_teeth)
    except Exception as e:
        print(f"❌ 解析 Probing Depth 失敗: {e}")
        sys.exit(1)

    ordered_groups = [
        [18, 17, 16, 15, 14, 13, 12, 11],
        [21, 22, 23, 24, 25, 26, 27, 28],
        [38, 37, 36, 35, 34, 33, 32, 31],
        [41, 42, 43, 44, 45, 46, 47, 48],
    ]


    print("LF:")
    print(generate_present_dentition(df_raw, tooth_rows_idx, missing_teeth))
    print("") 
    
    print("2. Full mouth general calculus and plaque deposition")
    print("")
    
    pd_report_str = get_flagged_teeth_string(records, ordered_groups, gap=6)
    print(pd_report_str.rstrip()) 
    print("")

    try:
        print_mobility_summary(csv_path, missing_teeth)
    except Exception as e:
        print("4.Mobility: nil")

    try:
        print_furcation_summary(csv_path, missing_teeth)
    except Exception as e:
        print("\n5.Furcation: nil")
        
    print("\n" + "="*60 + "\n")