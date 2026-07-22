# app.py
"""
Streamlit 主程式：自動判斷數據類型，動態切換單期 (6 頁) 或雙期對比 (6+6 頁) 簡報下載載點。
"""
import streamlit as st
from core_parser import parse_periodontal_csv
from pptx_engine import create_six_sextants_presentation, create_comparison_presentation

# 頁面配置
st.set_page_config(page_title="AI 牙周六象限簡報直出系統", layout="wide")

st.title("🦷 牙周檢查表六象限簡報直出系統")
st.subheader("Director 臨床與研討會專用版")

# 檔案上傳器
uploaded_file = st.file_uploader("請上傳牙周檢查表 CSV / Excel 檔案", type=["csv", "xlsx"])

if uploaded_file is not None:
    with st.spinner("後端解析引擎處理中，請稍候..."):
        # 呼叫解析引擎 (已無需計算臨床病歷摘要)
        records, missing_teeth, is_comparison, patient_info = parse_periodontal_csv(uploaded_file)

    st.markdown("### 📊 六象限原生醫學簡報直出")
    
    # 呈現病患資訊 (若有解析出資訊)
    if patient_info.get('name') or patient_info.get('case_no'):
        st.caption(f"👤 病患名稱：{patient_info.get('name', '未填寫')} | 🪪 病歷號：{patient_info.get('case_no', '未填寫')}")

    # 1. 必備載點：Initial 6 象限簡報
    ppt_initial = create_six_sextants_presentation(records, missing_teeth)
    st.download_button(
        label="🛸 下載【Initial 六象限簡報】(6 頁 PPTX)",
        data=ppt_initial,
        file_name="Periodontal_Sextants_Initial.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="btn_initial",
        use_container_width=True
    )

    st.divider()

    # 2. 條件載點：Initial vs Re-evaluation 雙期對比簡報 (根據檔案結構動態解鎖)
    if is_comparison:
        st.success("✨ 檢測到 Re-evaluation 對比數據！已解鎖雙期對比簡報下載。")
        ppt_comparison = create_comparison_presentation(records, missing_teeth)
        st.download_button(
            label="🛸 下載【Initial vs Re-eval 對比簡報】(6+6 頁 PPTX)",
            data=ppt_comparison,
            file_name="Periodontal_Sextants_Initial_vs_ReEval.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="btn_comparison",
            use_container_width=True
        )
    else:
        st.info("💡 目前檔案為單期數據。若上傳包含 Re-evaluation 的 CSV，將自動開啟【6+6 頁雙期對比簡報】載點。")
