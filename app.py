import streamlit as st
import re
import pandas as pd
import tempfile
import os
import zipfile
import fitz  # PyMuPDF
from docx import Document

# --- 1. 页面配置 ---
st.set_page_config(page_title="简历筛选系统", layout="wide")
st.title("📄 简历筛选系统")

# --- 2. 初始化 Session State ---
if 'rules' not in st.session_state:
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
    
    uploaded_zip = st.file_uploader(
        "📦 上传简历文件夹 (请压缩为 ZIP 格式)",
        type=["zip"],
        help="支持包含子文件夹的 ZIP 包。系统会自动解压并扫描所有 PDF/DOCX 文件。"
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
            if re.search(r'\b' + re.escape(keyword) + r'\b', text, re.IGNORECASE):
                total_score += weight
                details[keyword] = weight
            else:
                details[keyword] = 0
                
        return total_score, details

    # --- 6. 触发筛选按钮 ---
    if st.button("🎯 开始智能筛选", use_container_width=True, type="primary"):
        if not uploaded_zip:
            st.warning("请先上传 ZIP 文件！")
        elif not st.session_state.rules:
            st.warning("请先在左侧添加至少一个筛选规则！")
        else:
            results = []
            # 创建临时目录
            with tempfile.TemporaryDirectory() as tmp_dir:
                st.info(f"📂 正在解压 ZIP 文件到临时目录...")
                # 解压 ZIP
                zip_path = os.path.join(tmp_dir, "uploaded.zip")
                with open(zip_path, "wb") as f:
                    f.write(uploaded_zip.getvalue())
                
                extract_dir = os.path.join(tmp_dir, "extracted")
                os.makedirs(extract_dir, exist_ok=True)
                
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                
                # 递归查找所有 PDF/DOCX 文件
                all_files = []
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.lower().endswith(('.pdf', '.docx')):
                            all_files.append(os.path.join(root, file))
                
                if not all_files:
                    st.error("
