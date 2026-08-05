"""
Streamlit 主應用程式入口與模組路由 (Page Router)
"""
import os
import re
import io
import streamlit as st
import pandas as pd
from contextlib import redirect_stdout

import config
from core_parser import (
    find_missing_teeth, 
    is_comparison_file,
    find_tooth_rows,
    find_missing_rows,
    get_missing_teeth_set,
    parse_periodontal_pd,
    generate_present_dentition,
    get_flagged_teeth_string,
    print_mobility_summary,
    print_furcation_summary
)
from pptx_engine import create_six_sextants_presentation, create_comparison_presentation

# 🚀 匯入獨立的 12 口內照處理模組
from intraoral_photos import render_intraoral_photo_page

st.set_page_config(
    page_title="Periodontal & Intraoral Photo Generator",
    page_icon="🦷",
    layout="wide"
)

# ============================================================
# 輔助函式：動態轉換下載檔名
# ============================================================
def format_download_filename(original_filename: str, is_comparison: bool) -> str:
    """
    解析上傳 CSV 的原始檔名，自動轉化為標準簡報檔名格式。
    """
    base_name = os.path.splitext(original_filename)[0]
    clean_name = re.sub(r'\([^\)]*\)', '', base_name).strip()
    match = re.search(r'([^\d_\s]+)[_\s]*(\d+)', clean_name)
    
    if match:
        name = match.group(1).strip()
        patient_id = match.group(2).strip()
        patient_info = f"{name}{patient_id}"
    else:
        patient_info = "Peri_Report"

    clean_name_lower = clean_name.lower()
    
    if is_comparison or "initial&re-evaluation" in clean_name_lower or "initial & re-evaluation" in clean_name_lower:
        suffix = "I&Re"
    elif "re-evaluation" in clean_name_lower or "re" in clean_name_lower:
        suffix = "Re"
    elif "initial" in clean_name_lower:
        suffix = "I"
    else:
        suffix = "Report"

    return f"{patient_info}_{suffix}.pptx"

# ============================================================
# 輔助函式：生成 Objective 病歷文字
# ============================================================
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
# ============================================================
# 頂部大標題區塊
# ============================================================
st.title("🦷 PerioGeneratorPro")
st.markdown('<div style="text-align: right;"><b>【 Established by B09 王冠中 】</b></div>', unsafe_allow_html=True)

# 步驟 1 與 步驟 2
st.markdown("1. 去你要下載的 Charting Google sheet")
st.markdown("2. 選擇對的工作簿（Initial, Re-evaluation, or Initial & Re-evaluation)")

# 🚀 橫跨整頁顯示 Labels.png
image_path_labels = "Labels.png"
if os.path.exists(image_path_labels):
    st.image(image_path_labels, use_container_width=True)
else:
    st.warning(f"⚠️ 找不到說明圖片：{image_path_labels}")

# 步驟 3 至 步驟 6
st.markdown("3. 點擊左上方「檔案」")
st.markdown("4. 點擊「下載」")
st.markdown("5. 下載「逗號分隔值檔案（.csv）」")

col1, col2 = st.columns([3, 4])  # 左邊佔 3/7，右邊佔 4/7
with col1:
    image_path_download = "Download.png"
    if os.path.exists(image_path_download):
        st.image(image_path_download, use_container_width=True)
    else:
        st.warning(f"⚠️ 找不到說明圖片：{image_path_download}")

st.markdown("6. 去電腦的下載項目資料夾，找到要上傳的檔案！")

st.write("---")
st.write("")

def render_periodontal_generator_page():
    uploaded_file = st.file_uploader("請上傳要轉換的 CSV 檔案", type=["csv"])

    if uploaded_file is not None:
        try:
            # 讀取 CSV (強制字串讀取，避免數字型別誤判)
            df = pd.read_csv(uploaded_file, header=None, dtype=str).fillna("")
            
            missing_teeth = find_missing_teeth(df)
            is_comparison = is_comparison_file(df)
            output_filename = format_download_filename(uploaded_file.name, is_comparison)

            # ----------------------------------------------------
            # 1. 簡報生成與下載區塊
            # ----------------------------------------------------
            if is_comparison:
                st.success("生成 Initial & Re-evaluation 簡報，請稍候")
                ppt_comparison = create_comparison_presentation(df, missing_teeth)
                st.download_button(
                    label="📥 下載簡報",
                    data=ppt_comparison,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                file_name_lower = uploaded_file.name.lower()
                if "re-evaluation" in file_name_lower or "re" in file_name_lower:
                    stage_name = "Re-evaluation"
                else:
                    stage_name = "Initial"

                st.info(f"生成 {stage_name} 簡報，請稍候")
                ppt_initial = create_six_sextants_presentation(df, missing_teeth)
                st.download_button(
                    label="📥 下載簡報",
                    data=ppt_initial,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            st.write("---")

            # ----------------------------------------------------
            # 2. Objective 病歷紀錄文字區塊 (位在簡報下載與免責聲明之間)
            # ----------------------------------------------------
            with st.container(border=True):
                st.subheader("📋 Objective 病歷紀錄文字")
                
                # 自動生成文字內容
                objective_text = generate_objective_text(df)
                
                # 提示與可直接複製的程式碼框
                st.caption("點擊下方右上角按鈕即可快速複製文字貼至電子病歷：")
                st.code(objective_text, language="text")

            st.write("---")

        except Exception as e:
            st.error(f"解析檔案時發生錯誤：{str(e)}")

def main():
    st.sidebar.title("🦷")
    page_choice = st.sidebar.radio(
        "請選擇功能模組：",
        ["牙周簡報生成 (PPT)", "12 口內照上傳展示 (Photos)"]
    )

    if page_choice == "牙周簡報生成 (PPT)":
        render_periodontal_generator_page()
        
        # ============================================================
        # 區塊：隱私保護與使用免責聲明 
        # ============================================================
        with st.expander("⚖️ 隱私保護與免責聲明", expanded=False):
            st.markdown("""
                <style>
                .disclaimer-text {
                    font-size: 13px;
                    color: #666666;
                    line-height: 1.6;
                }
                </style>
                <div class="disclaimer-text">
                    <ul>
                        <li><b>零資料留存</b>：本工具為純前端即時解析軟體，原始碼絕不具備任何資料庫上傳、硬碟儲存或歷史紀錄留存功能。</li>
                        <li><b>即時銷毀</b>：上傳之 CSV 檔案僅短暫載入伺服器記憶體中進行排版計算，網頁一旦關閉或重新整理，數據立刻完全銷毀。</li>
                        <li><b>去識別化建議</b>：強烈建議使用者在上傳前，先將檔案內之病患敏感資訊（如真實姓名、身分證字號）匿名、或是去識別化。</li>
                    </ul>
                </div>
            """, unsafe_allow_html=True)

    elif page_choice == "12 口內照上傳展示 (Photos)":
        render_intraoral_photo_page()
      
if __name__ == "__main__":
    main()
