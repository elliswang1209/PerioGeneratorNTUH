# app.py
"""
Streamlit 主程式：根據上傳數據類型，單一動態切換下載載點。
- 雙期檔：僅提供【Initial vs Re-eval 對比簡報】(6+6 頁 PPTX)
- 單期檔：僅提供【Initial 六象限簡報】(6 頁 PPTX)
"""
import streamlit as st
from core_parser import parse_periodontal_csv
from pptx_engine import create_six_sextants_presentation, create_comparison_presentation

# 頁面配置
st.set_page_config(page_title="AI 牙周六象限簡報直出系統", layout="wide")

st.title("🦷 PerioGeneratorPro")
st.subheader("可以生成 Initial 或是 Initial & Re-evalaution 的對比")

# 檔案上傳器
uploaded_file = st.file_uploader("請上傳牙周檢查表 CSV / Excel 檔案", type=["csv", "xlsx"])

if uploaded_file is not None:
    with st.spinner("處理中，請稍候..."):
        # 呼叫解析引擎
        df, missing_teeth, is_comparison, patient_info = parse_periodontal_csv(uploaded_file)

    st.markdown("### 📊 簡報")
    
    # 呈現病患資訊 (若有解析出資訊)
    if patient_info.get('name') or patient_info.get('case_no'):
        st.caption(f"{patient_info.get('name', '未填寫')} | {patient_info.get('case_no', '未填寫')}")

    # 🚀 根據數據類型僅顯示「單一」對應下載載點
    if is_comparison:
        st.success("成功生成！")
        ppt_comparison = create_comparison_presentation(df, missing_teeth)
        st.download_button(
            label="下載 Initial vs Re-eval 對比簡報",
            data=ppt_comparison,
            file_name="Periodontal_Sextants_Initial_vs_ReEval.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="btn_comparison",
            use_container_width=True
        )
    else:
        st.info("💡 檢測到 Initial 單期數據。已為您生成 6 頁六象限簡報。")
        ppt_initial = create_six_sextants_presentation(df, missing_teeth)
        st.download_button(
            label="🛸 下載【Initial 六象限簡報】(6 頁 PPTX)",
            data=ppt_initial,
            file_name="Periodontal_Sextants_Initial.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="btn_initial",
            use_container_width=True
        )
