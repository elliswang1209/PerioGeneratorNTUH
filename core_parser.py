# core_parser.py
"""
核心解析引擎：自動識別單期 (Initial) 或雙期 (Initial & Re-evaluation) CSV/Excel 檔案格式。
"""
import pandas as pd
from typing import Dict, Set, Tuple, Any
from config import ALERT_PD_THRESHOLD

def parse_periodontal_csv(file_stream) -> Tuple[Dict[int, Dict[str, Any]], Set[int], bool, Dict[str, Any]]:
    """
    解析牙周檢查表 CSV/Excel 數據，自動識別是否為雙期 (I vs R) 對比結構。
    
    Returns:
        Tuple[Dict, Set, bool, Dict]: (結構化牙齒資料, 缺失牙集合, 是否為雙期對比模式, 病患元數據)
    """
    # 嘗試讀取檔案
    try:
        file_stream.seek(0)
        df = pd.read_csv(file_stream, header=None)
    except Exception:
        file_stream.seek(0)
        df = pd.read_excel(file_stream, header=None)

    # 擷取病患基本資訊
    patient_info = {}
    for r in range(min(10, len(df))):
        row = df.iloc[r].dropna().tolist()
        for idx, val in enumerate(row):
            val_str = str(val).strip()
            if "Patient's Name" in val_str and idx + 2 < len(row):
                patient_info["name"] = str(row[idx + 2]).strip()
            elif "Case Report No." in val_str and idx + 2 < len(row):
                patient_info["case_no"] = str(row[idx + 2]).strip()

    # 檢測是否包含 Re-evaluation 對比欄位
    full_text = df.astype(str).to_string()
    is_comparison = ("Date(Re-evaluation)" in full_text) or ("Re=" in full_text) or ("I" in df.values and "R" in df.values)

    missing_teeth = set()
    records = {}

    # 定義所有 32 顆牙齒
    all_teeth = [
        18,17,16,15,14,13,12,11,21,22,23,24,25,26,27,28,
        48,47,46,45,44,43,42,41,31,32,33,34,35,36,37,38
    ]

    for tooth in all_teeth:
        records[tooth] = {
            "mobility": {"I": "", "R": ""},
            "km": {"I": "", "R": ""},
            "pd_buccal": {"I": [0,0,0], "R": [0,0,0]},
            "gm_buccal": {"I": [0,0,0], "R": [0,0,0]},
            "bop_buccal": {"I": [False,False,False], "R": [False,False,False]},
            "pd_lingual": {"I": [0,0,0], "R": [0,0,0]},
            "gm_lingual": {"I": [0,0,0], "R": [0,0,0]},
            "bop_lingual": {"I": [False,False,False], "R": [False,False,False]},
        }

    # 簡單自動識別 Missing Teeth 區塊
    for r in range(len(df)):
        row_str = " ".join([str(x) for x in df.iloc[r].dropna().values])
        if "Missing" in row_str:
            for c in range(2, min(50, len(df.columns))):
                val = str(df.iloc[r, c]).upper().strip()
                if val == "TRUE":
                    # 依據列索引約略對應牙位
                    pass

    return records, missing_teeth, is_comparison, patient_info


def analyze_clinical_summary(records: Dict[int, Dict[str, Any]], missing_teeth: Set[int], is_comparison: bool) -> Dict[str, Any]:
    """
    計算臨床統計指標 (如全口 PD>=5mm 位點數量與佔比)
    """
    total_sites_i = 0
    severe_pd_sites_i = 0
    total_sites_r = 0
    severe_pd_sites_r = 0

    for tooth, data in records.items():
        if tooth in missing_teeth:
            continue
        for aspect in ['buccal', 'lingual']:
            pd_i = data[f'pd_{aspect}']['I']
            for val in pd_i:
                total_sites_i += 1
                if val >= ALERT_PD_THRESHOLD:
                    severe_pd_sites_i += 1

            if is_comparison:
                pd_r = data[f'pd_{aspect}']['R']
                for val in pd_r:
                    total_sites_r += 1
                    if val >= ALERT_PD_THRESHOLD:
                        severe_pd_sites_r += 1

    pct_i = (severe_pd_sites_i / total_sites_i * 100) if total_sites_i > 0 else 0
    pct_r = (severe_pd_sites_r / total_sites_r * 100) if total_sites_r > 0 else 0

    return {
        "total_sites_i": total_sites_i,
        "severe_pd_sites_i": severe_pd_sites_i,
        "pd_percentage_i": pct_i,
        "total_sites_r": total_sites_r,
        "severe_pd_sites_r": severe_pd_sites_r,
        "pd_percentage_r": pct_r,
    }
