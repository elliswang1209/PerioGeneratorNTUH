"""
Streamlit 主應用程式入口與模組路由 (Page Router)
"""
import os
import re
import streamlit as st
import pandas as pd

import config
from core_parser import find_missing_teeth, is_comparison_file
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
    範例：
    - 陳雅秋_6707208 - Initial(輸入).csv -> 陳雅秋6707208_I.pptx
    - 劉頂立_8066684 - Re-evaluation(輸入).csv -> 劉頂立8066684_Re.pptx
    - 蔡怡文_7214894 - Initial&Re-evaluation(輸出).csv -> 蔡怡文7214894_I&Re.pptx
    """
    # 1. 移除副檔名 .csv
    base_name = os.path.splitext(original_filename)[0]
    
    # 2. 清理檔名尾端的 (輸入)、(輸出) 等括弧註記
    clean_name = re.sub(r'\([^\)]*\)', '', base_name).strip()
    
    # 3. 使用正則表達式提取「姓名」與「病歷號數字」
    # 匹配模式：非數字姓名 + 分隔符(底線或空格) + 純數字病歷號
    match = re.search(r'([^\d_\s]+)[_\s]*(\d+)', clean_name)
    
    if match:
        name = match.group(1).strip()
        patient_id = match.group(2).strip()
        patient_info = f"{name}{patient_id}"
    else:
        patient_info = "Peri_Report"

    # 4. 判斷階段標籤 (I&Re / Re / I)
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
# 頂部大標題區塊
# ============================================================
st.title("🦷 PerioGenerator")
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
            df = pd.read_csv(uploaded_file, header=None)
            
            missing_teeth = find_missing_teeth(df)
            is_comparison = is_comparison_file(df)
            output_filename = format_download_filename(uploaded_file.name, is_comparison)

            if is_comparison:
                st.success("生成 Initial & Re-evaluation 簡報，請稍候")
                ppt_comparison = create_comparison_presentation(df, missing_teeth)
                st.download_button(
                    label="📥 下載簡報",
                    data=ppt_comparison,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.write("---")
                st.write("")
            else:
                # 🚀 動態判斷是 Initial 還是 Re-evaluation
                file_name_lower = uploaded_file.name.lower()
                if "re-evaluation" in file_name_lower or "re" in file_name_lower:
                    stage_name = "Re-evaluation"
                else:
                    stage_name = "Initial"

                st.info(f"生成 {stage_name} 簡報，請稍候")
                ppt_initial = create_six_sextants_presentation(df, missing_teeth)
                st.download_button(
                    label=f"📥 下載簡報",
                    data=ppt_initial,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.write("---")
                st.write("")


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
        # 區塊 1：隱私保護與使用免責聲明 
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
