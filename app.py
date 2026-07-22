# app.py
"""
Streamlit 主程式：自動判斷數據類型，動態切換單期 (6 頁) 或雙期對比 (6+6 頁) 下載載點。
"""
import streamlit as st
import streamlit.components.v1 as components
from core_parser import parse_periodontal_csv, analyze_clinical_summary
from pptx_engine import create_six_sextants_presentation, create_comparison_presentation

# 頁面配置
st.set_page_config(page_title="AI 牙周病歷自動化與六象限簡報直出系統", layout="wide")

st.title("🦷 牙周檢查表自動化分析與六象限簡報直出系統")
st.subheader("Director 臨床與研討會專用版")

# 檔案上傳
uploaded_file = st.file_uploader("請上傳牙周檢查表 CSV / Excel 檔案", type=["csv", "xlsx"])

if uploaded_file is not None:
    with st.spinner("後端解析引擎處理中，請稍候..."):
        records, missing_teeth, is_comparison, patient_info = parse_periodontal_csv(uploaded_file)
        summary = analyze_clinical_summary(records, missing_teeth, is_comparison)

    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("### 📋 自動生成臨床病歷文字")

        report_i = f"""【臨床牙周客觀檢查摘要 - Initial】
病患名稱：{patient_info.get('name', '未填寫')} | 病歷號：{patient_info.get('case_no', '未填寫')}
全口探測點位共 {summary['total_sites_i']} 點。
探測深度 PD >= 5mm 之嚴重病變點位計 {summary['severe_pd_sites_i']} 點 (佔全口 {summary['pd_percentage_i']:.1f}%)。
        """

        if is_comparison:
            report_r = f"""
【臨床牙周客觀檢查摘要 - Re-evaluation 對比】
再評估 PD >= 5mm 之嚴重位點計 {summary['severe_pd_sites_r']} 點 (佔全口 {summary['pd_percentage_r']:.1f}%)。
病變位點改善率：{((summary['severe_pd_sites_i'] - summary['severe_pd_sites_r']) / max(1, summary['severe_pd_sites_i']) * 100):.1f}%
            """
            medical_report = report_i + report_r
        else:
            medical_report = report_i

        st.code(medical_report, language="text")

        # JavaScript 一鍵複製病歷文字區塊
        js_copy_html = """
        <button id="copy-btn" style="
            background-color: #1E88E5; color: white; border: none; 
            padding: 8px 16px; border-radius: 4px; cursor: pointer; font-weight: bold;">
            📋 一鍵複製病歷文字
        </button>
        <script>
            document.getElementById('copy-btn').addEventListener('click', function() {
                const codeBlocks = window.parent.document.querySelectorAll('code');
                if(codeBlocks.length > 0) {
                    const textToCopy = codeBlocks[0].innerText;
                    window.parent.navigator.clipboard.writeText(textToCopy).then(() => {
                        alert('🎉 病歷文字已成功複製至剪貼簿！');
                    }).catch(err => {
                        alert('複製失敗，請手動選取複製。');
                    });
                }
            });
        </script>
        """
        components.html(js_copy_html, height=50)

    with col2:
        st.markdown("### 📊 六象限原生醫學簡報直出")

        # 1. 第一個載點：Initial 6 象限簡報 (必備)
        ppt_initial = create_six_sextants_presentation(records, missing_teeth)
        st.download_button(
            label="🛸 下載【Initial 六象限簡報】(6 頁 PPTX)",
            data=ppt_initial,
            file_name="Periodontal_Sextants_Initial.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="btn_initial"
        )

        st.divider()

        # 2. 第二個載點：Initial vs Re-evaluation 雙期對比簡報 (根據檔案結構動態解鎖)
        if is_comparison:
            st.success("✨ 檢測到 Re-evaluation 對比數據！已解鎖雙期對比簡報下載。")
            ppt_comparison = create_comparison_presentation(records, missing_teeth)
            st.download_button(
                label="🛸 下載【Initial vs Re-eval 對比簡報】(6+6 頁 PPTX)",
                data=ppt_comparison,
                file_name="Periodontal_Sextants_Initial_vs_ReEval.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="btn_comparison"
            )
        else:
            st.info("💡 目前檔案為單期數據。若上傳包含 Re-evaluation 的 CSV，將自動開啟【6+6 頁雙期對比簡報】載點。")
