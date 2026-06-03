import streamlit as st
import re
import pandas as pd
import tempfile
import os
import fitz  # PyMuPDF
from docx import Document

# --- 1. 页面配置 ---
st.set_page_config(page_title="光学工程师简历筛选系统", layout="wide")
st.title("📄 光学工程师简历筛选系统")

# --- 2. 初始化 Session State ---
if 'rules' not in st.session_state:
    # 注意：这里已经把“光学设计”去掉了，如果你想加回来，可以在网页上手动加
    st.session_state.rules = [
        {'keyword': 'Zemax', 'weight': 8, 'desc': '熟练使用 Zemax/CodeV 进行仿真'},
        {'keyword': 'Python', 'weight': 5, 'desc': '具备编程开发能力'},
    ]

# --- 3. UI 布局：左右分栏 ---
col_setup, col_result = st.columns([1, 2])

with col_setup:
    st.header("⚙️ 自定义筛选规则")
    
    with st.expander("点击展开：设置评分标准", expanded=True):
        
        # 显示现有的规则表格
        st.subheader("当前评分规则")
        if st.session_state.rules:
            rule_df_data = []
            for i, rule in enumerate(st.session_state.rules):
                rule_df_data.append({
                    "序号": i + 1,
                    "筛选关键词": rule['keyword'],
                    "权重分": rule['weight'],
                    "备注说明": rule['desc']
                })
            st.dataframe(pd.DataFrame(rule_df_data), use_container_width=True, hide_index=True)
        else:
            st.info("暂无规则，请在下方添加。")

        st.markdown("---")
        
        # --- 新增/修改/删除逻辑 ---
        st.subheader("✍️ 新增或修改规则")
        
        input_keyword = st.text_input("输入技能关键词 (如：激光雷达)", key="input_kw")
        input_weight = st.number_input("设置权重分", min_value=1, max_value=100, value=10, step=1, key="input_w")
        input_desc = st.text_input("备注说明 (可选)", key="input_d")

        # 按钮区域
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            if st.button("➕ 保存/更新规则", use_container_width=True, type="primary"):
                if not input_keyword:
                    st.warning("⚠️ 请输入关键词！")
                else:
                    # 核心逻辑：检查是否已存在
                    existing_index = -1
                    for i, rule in enumerate(st.session_state.rules):
                        if rule['keyword'].lower() == input_keyword.lower():
                            existing_index = i
                            break
                    
                    if existing_index != -1:
                        # 如果存在，先删除旧的
                        del st.session_state.rules[existing_index]
                        st.success(f"✅ 规则 '{input_keyword}' 已更新！")
                    else:
                        st.success(f"✅ 成功添加 '{input_keyword}'！")
                    
                    # 追加新的规则
                    st.session_state.rules.append({
                        "keyword": input_keyword,
                        "weight": input_weight,
                        "desc": input_desc
                    })
                    st.rerun()

        with btn_col2:
            if st.button("🗑️ 重置所有规则", use_container_width=True):
                st.session_state.rules = []
                st.rerun()

with col_result:
    st.header("🚀 开始筛选")
    
    uploaded_files = st.file_uploader(
        "上传简历 (支持多选)", 
        type=["pdf", "docx"], 
        accept_multiple_files=True,
        help="支持 PDF 和 Word 格式"
    )

    # --- 4. 核心逻辑：提取文本 ---
    def extract_text(file_path, file_name):
        text = ""
        try:
            if file_name.endswith('.pdf'):
                doc = fitz.open(file_path)
                for page in doc:
                    text += page.get_text()
            elif file_name.endswith('.docx'):
                doc = Document(file_path)
                for para in doc.paragraphs:
                    text += para.text + "\n"
        except Exception as e:
            st.error(f"读取文件 {file_name} 失败: {e}")
        return text

    # --- 5. 核心逻辑：计算分数 ---
    def calc_score(text, rules):
        total_score = 0
        details = {}
        
        for rule in rules:
            keyword = rule['keyword']
            weight = rule['weight']
            # 使用正则表达式进行不区分大小写的匹配
            if re.search(r'\b' + re.escape(keyword) + r'\b', text, re.IGNORECASE):
                total_score += weight
                details[keyword] = weight
            else:
                details[keyword] = 0
                
        return total_score, details

    # --- 6. 触发筛选按钮 ---
    if st.button("🎯 开始智能筛选", use_container_width=True, type="primary"):
        if not uploaded_files:
            st.warning("请先上传简历！")
        elif not st.session_state.rules:
            st.warning("请先在左侧添加至少一个筛选规则！")
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
            
            st.dataframe(
                df.sort_values(by="💯 匹配总分", ascending=False),
                use_container_width=True,
                hide_index=True
            )
            
            csv = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("⬇️ 下载筛选报告 (CSV)", csv, "简历筛选报告.csv", "text/csv")
