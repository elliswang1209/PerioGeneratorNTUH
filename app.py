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
def generate_objective_text(df_raw: pd.DataFrame) -> str:
    """
    解析 CSV 原始資料，編織生成標準的 Objective 病歷紀錄文字。
    """
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
        # 1. Present dentition
        print(generate_present_dentition(df_raw, tooth_rows_idx, missing_teeth))
        print("")
        
        # 2. Oral hygiene
        print(" 2. Oral hygiene: ______; generalized gingival inflammation with plaque and calculus deposition.")
        print("")
        
        # 3. Probing depth summary
        pd_report_str = get_flagged_teeth_string(records, ordered_groups, gap=6)
        print(pd_report_str.rstrip())
        print("")
        
        # 4. Mobility summary
        try:
            print_mobility_summary(df_raw, missing_teeth)
        except Exception:
            print(" 4. Mobility: nil")
        
        # 5. Furcation summary
        try:
            print_furcation_summary(df_raw, missing_teeth)
        except Exception:
            print("\n 5. Furcation: nil")

    return output_buffer.getvalue()

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
