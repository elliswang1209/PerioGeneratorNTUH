"""
Streamlit 主應用程式入口與模組路由 (Page Router)
"""
import os
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
   
    
    uploaded_file = st.file_uploader("請選擇 Periogrid 輸出的 CSV 檔案", type=["csv"])

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file, header=None)
            st.success("CSV 檔案讀取成功！")
            
            missing_teeth = find_missing_teeth(df)
            is_comparison = is_comparison_file(df)

            st.markdown("### 生成簡報下載")
            if is_comparison:
                st.info("此 CSV 包含 Initial & Re-evaluation")
                ppt_comparison = create_comparison_presentation(df, missing_teeth)
                st.download_button(
                    label="📥 下載 Initial & Re-evaluation 對比簡報",
                    data=ppt_comparison,
                    file_name="Peri_CC_Report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                # 增加一點垂直間距
                st.write("---")
                st.write("")
            else:
                st.info("此 CSV 為 Initial or Re-evaluation")
                ppt_initial = create_six_sextants_presentation(df, missing_teeth)
                st.download_button(
                    label="📥 下載 Initial 簡報",
                    data=ppt_initial,
                    file_name="Peri_initial_Report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                # 增加一點垂直間距
                st.write("---")
                st.write("")
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
    elif page_choice == "12 口內照上傳展示 (Photos)":
        render_intraoral_photo_page()
     
if __name__ == "__main__":
    main()
