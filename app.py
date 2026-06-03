import streamlit as st
import re
import pandas as pd
import tempfile
import os
import fitz  # PyMuPDF
from docx import Document

st.set_page_config(page_title="光学工程师简历筛选系统", layout="wide")
st.title("📄 光学工程师简历筛选系统")

if 'rules' not in st.session_state:
    st.session_state.rules = [
        {'keyword': '光学设计', 'weight': 10, 'desc': '掌握成像光学/照明系统设计'},
        {'keyword': 'Zemax', 'weight': 8, 'desc': '熟练使用 Zemax/CodeV 进行仿真'},
        {'keyword': '项目管理', 'weight': 5, 'desc': '具备项目统筹或跨部门协作经验'},
    ]

col_setup, col_result = st.columns([1, 2])

with col_setup:
    st.header("⚙️ 自定义筛选规则")
    with st.expander("点击展开：设置评分标准"):
        st.subheader("当前评分规则")
        rule_df_data = []
        for i, rule in enumerate(st.session_state.rules):
            rule_df_data.append({
                "序号": i + 1,
                "筛选关键词": rule['keyword'],
                "权重分": rule['weight'],
                "备注说明": rule['desc']
            })
        st.dataframe(pd.DataFrame(rule_df_data), use_container_width=True, hide_index=True)
        st.markdown("---")
        st.subheader("✍️ 新增/修改规则")
        with st.form("add_rule_form"):
            new_keyword = st.text_input("输入技能关键词 (如：激光雷达)")
            new_weight = st.number_input("设置权重分", min_value=1, max_value=100, value=10)
            new_desc = st.text_input("备注说明 (可选)")
            submitted = st.form_submit_button("➕ 添加到规则表")
            if submitted and new_keyword:
                exists = False
                for rule in st.session_state.rules:
                    if rule['keyword'].lower() == new_keyword.lower():
                        rule['weight'] = new_weight
                        rule['desc'] = new_desc
                        exists = True
                        break
                if not exists:
                    st.session_state.rules.append({
                        'keyword': new_keyword,
                        'weight': new_weight,
                        'desc': new_desc
                    })
                st.rerun()

with col_result:
    st.header("🚀 开始筛选")
    uploaded_files = st.file_uploader("上传简历 (支持多选)", type=["pdf", "docx"], accept_multiple_files=True)
    def extract_text(file_path, file_name):
        text = ""
        try:
            if file_name.endswith('.pdf'):
                doc = fitz.open(file_path)
                for page in doc: text += page.get_text()
            elif file_name.endswith('.docx'):
                doc = Document(file_path)
                for para in doc.paragraphs: text += para.text + "\n"
        except: pass
        return text
    def calc_score(text, rules):
        total_score = 0
        details = {}
        for rule in rules:
            keyword = rule['keyword']
            weight = rule['weight']
            if re.search(r'\b' + re.escape(keyword) + r'\b', text, re.IGNORECASE):
                total_score += weight
                details[keyword] = weight
            else:
                details[keyword] = 0
        return total_score, details
    if st.button("🎯 开始智能筛选", use_container_width=True, type="primary"):
        if not uploaded_files:
            st.warning("请先上传简历！")
        else:
            results = []
            progress_bar = st.progress(0, text="正在解析简历...")
            for i, uf in enumerate(uploaded_files):
                suffix = os.path.splitext(uf.name)[1]
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(uf.getvalue())
                    tmp_path = tmp.name
                text = extract_text(tmp_path, uf.name)
                total, detail = calc_score(text, st.session_state.rules)
                results.append({
                    "👤 候选人": uf.name.replace('.pdf', '').replace('.docx', ''),
                    "💯 匹配总分": total,
                    **detail
                })
                os.unlink(tmp_path)
                progress_bar.progress((i + 1) / len(uploaded_files), text=f"正在分析: {uf.name}")
            progress_bar.empty()
            st.success(f"筛选完成，共处理 {len(results)} 份简历")
            df = pd.DataFrame(results)
            cols = df.columns.tolist()
            score_cols = [c for c in cols if c not in ["👤 候选人", "💯 匹配总分"]]
            df = df[["👤 候选人", "💯 匹配总分"] + score_cols]
            st.dataframe(df.sort_values(by="💯 匹配总分", ascending=False), use_container_width=True, hide_index=True)
            csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("⬇️ 下载筛选报告 (CSV)", csv, "简历筛选报告.csv", "text/csv")
