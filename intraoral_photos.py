# intraoral_photos.py
"""
12 張口內照上傳與展示處理模組
"""
import streamlit as st
from PIL import Image
import os

# 定義標準 12 張口內照項目
PHOTO_LABELS = [
    "1. Upper Arch (上顎咬合面)",
    "2. Lower Arch (下顎咬合面)",
    "3. Anterior Frontal View (前牙正觀)",
    "4. Right Lateral View (右側咬合觀)",
    "5. Left Lateral View (左側咬合觀)",
    "6. Upper Right Posterior (上顎右側後牙)",
    "7. Upper Anterior Palatal (上顎前牙腭側)",
    "8. Upper Left Posterior (上顎左側後牙)",
    "9. Lower Right Posterior (下顎右側後牙)",
    "10. Lower Anterior Lingual (下顎前牙舌側)",
    "11. Lower Left Posterior (下顎左側後牙)",
    "12. Overjet / Overbite View (覆咬觀)"
]

def render_intraoral_photo_page():
    st.header("📸 12 張口內照上傳與圖片展示系統")
    st.caption("請上傳患者的 12 張標準口內照片，系統將自動進行版面編排與預覽展示。")

    st.markdown("---")

    # 初始化 Session State 儲存上傳圖片
    if "uploaded_photos" not in st.session_state:
        st.session_state["uploaded_photos"] = {}

    st.subheader("1. 分項照片上傳")
    
    # 使用 3 欄網格呈現上傳區域
    cols = st.columns(3)
    for idx, label in enumerate(PHOTO_LABELS):
        col_idx = idx % 3
        with cols[col_idx]:
            uploaded_file = st.file_uploader(
                label, 
                type=["jpg", "jpeg", "png"], 
                key=f"photo_upload_{idx}"
            )
            if uploaded_file is not None:
                image = Image.open(uploaded_file)
                st.session_state["uploaded_photos"][label] = image

    st.markdown("---")

    # 顯示上傳成果展示
    st.subheader("2. 口內照展示看板 (12-Grid Gallery)")
    
    uploaded_dict = st.session_state["uploaded_photos"]
    uploaded_count = len(uploaded_dict)
    
    st.info(f"目前已上傳：**{uploaded_count} / 12** 張照片")

    if uploaded_count > 0:
        # 展示網格 (每列 3 張)
        display_cols = st.columns(3)
        for idx, label in enumerate(PHOTO_LABELS):
            col_idx = idx % 3
            with display_cols[col_idx]:
                st.markdown(f"**{label}**")
                if label in uploaded_dict:
                    st.image(uploaded_dict[label], use_container_width=True)
                else:
                    # 未上傳時的預設占位顯示
                    st.warning("⚠️ 尚未上傳")
    else:
        st.write("請在上方區域選擇並上傳圖片，上傳後圖片將即時顯示於此展示牆。")
