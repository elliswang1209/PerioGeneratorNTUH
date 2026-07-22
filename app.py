# app.py
"""
Streamlit 主應用程式入口與模組路由 (Page Router)
"""
import streamlit as st
import pandas as pd

import config
from core_parser import find_missing_teeth
from pptx_engine import create_six_sextants_presentation, create_comparison_presentation

# 🚀 匯入獨立的 12 口內照處理模組
from intraoral_photos import render_intraoral_photo_page

st.set_page_config(
    page_title="Periodontal & Intraoral Photo Generator",
    page_icon="🦷",
    layout="wide"
)

def render_periodontal_generator_page():
    st.header("🦷 牙周病檢查表簡報生成器 (Periodontal Chart Generator)")
    st.caption("請上傳 Initial 及/或 Re-evaluation 的檢查 CSV 檔案，系統將自動解析並繪製高質感簡報。")

    st.markdown("---")
    
    uploaded_file = st.file_uploader("請選擇 Periogrid 輸出的 CSV 檔案", type=["csv"])

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file, header=None)
            st.success("CSV 檔案讀取成功！")
            
            missing_teeth = find_missing_teeth(df)
            
            col1, col2 = st.columns(2)
            with col1:
                st.write("**自動偵測缺牙牙號**：", sorted(list(missing_teeth)))

            # 判斷是否為 Initial & Re-evaluation 對比檔 (列數 > 50 粗略判斷)
            is_comparison = len(df) > 50 

            st.markdown("### 生成簡報下載")
            if is_comparison:
                st.info("偵測到此 CSV 包含 Initial & Re-evaluation 對比資料。")
                ppt_comparison = create_comparison_presentation(df, missing_teeth)
                st.download_button(
                    label="📥 下載 Comparison 對比簡報 (.pptx)",
                    data=ppt_comparison,
                    file_name="Periodontal_Comparison_Report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.info("偵測到此 CSV 為單一階段 Initial 資料。")
                ppt_initial = create_six_sextants_presentation(df, missing_teeth)
                st.download_button(
                    label="📥 下載 Initial 簡報 (.pptx)",
                    data=ppt_initial,
                    file_name="Periodontal_Initial_Report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

        except Exception as e:
            st.error(f"解析檔案時發生錯誤：{str(e)}")

def main():
    # 側邊欄導航選單
    st.sidebar.title("🦷 功能導航")
    page_choice = st.sidebar.radio(
        "請選擇功能模組：",
        ["牙周簡報生成 (PPT)", "12 口內照上傳展示 (Photos)"]
    )

    if page_choice == "牙周簡報生成 (PPT)":
        render_periodontal_generator_page()
    elif page_choice == "12 口內照上傳展示 (Photos)":
        # 🚀 導引呼叫獨立檔案 intraoral_photos.py 的頁面渲染函式
        render_intraoral_photo_page()

if __name__ == "__main__":
    main()
