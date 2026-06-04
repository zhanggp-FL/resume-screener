# ========== app.py (支持清除和修改默认值) ==========
import streamlit as st
import re
import pandas as pd
import tempfile
import os
import zipfile
import copy

import fitz
from docx import Document

# ---------- 1. 初始默认规则 (只读备份，用于初始化) ----------
INITIAL_DEFAULTS = {
    "光学设计工程师": {
        "基本技能": [{"keyword": "光学设计", "weight": 10, "desc": "几何/成像"}, {"keyword": "Zemax", "weight": 8, "desc": "软件"}],
        "专业技能": [{"keyword": "DOE", "weight": 6, "desc": "衍射"}, {"keyword": "杂散光", "weight": 5, "desc": "分析"}],
        "优选项": [{"keyword": "Python", "weight": 4, "desc": "自动化"}, {"keyword": "Matlab", "weight": 3, "desc": "算法"}]
    },
    "机械结构设计工程师": {
        "基本技能": [{"keyword": "SolidWorks", "weight": 10, "desc": "建模"}, {"keyword": "工程图", "weight": 8, "desc": "出图"}],
        "专业技能": [{"keyword": "公差分析", "weight": 7, "desc": "GD&T"}, {"keyword": "热设计", "weight": 5, "desc": "散热"}],
        "优选项": [{"keyword": "ANSYS", "weight": 4, "desc": "有限元"}, {"keyword": "项目管理", "weight": 3, "desc": "协作"}]
    },
    "CAE工程师": {
        "基本技能": [{"keyword": "Hypermesh", "weight": 10, "desc": "前处理"}, {"keyword": "Abaqus", "weight": 8, "desc": "求解"}],
        "专业技能": [{"keyword": "模态分析", "weight": 7, "desc": "动力学"}, {"keyword": "疲劳分析", "weight": 6, "desc": "寿命"}],
        "优选项": [{"keyword": "Python", "weight": 5, "desc": "二次开发"}, {"keyword": "二次开发", "weight": 4, "desc": "脚本"}]
    }
}

# ---------- 2. Session State 初始化 (核心改动) ----------
# 使用一个可变的字典来存储“当前生效的默认值”
if "DEFAULT_JOBS_DB" not in st.session_state:
    st.session_state.DEFAULT_JOBS_DB = copy.deepcopy(INITIAL_DEFAULTS)

if "current_job" not in st.session_state:
    st.session_state.current_job = list(st.session_state.DEFAULT_JOBS_DB.keys())[0]

if "rules" not in st.session_state:
    st.session_state.rules = copy.deepcopy(st.session_state.DEFAULT_JOBS_DB[st.session_state.current_job])

# ---------- 3. 页面配置 ----------
st.set_page_config(page_title="多岗位简历筛选系统", layout="wide")
st.title("📄 多岗位简历智能筛选系统")

left, right = st.columns([1, 2])

# ---------- 4. 左侧：岗位切换与规则管理 ----------
with left:
    st.header("⚙️ 岗位与规则配置")
    
    # --- 岗位选择器 ---
    job_options = list(st.session_state.DEFAULT_JOBS_DB.keys())
    current_idx = job_options.index(st.session_state.current_job)
    
    selected_job = st.selectbox("选择招聘岗位", options=job_options, index=current_idx)
    
    if selected_job != st.session_state.current_job:
        st.session_state.current_job = selected_job
        st.session_state.rules = copy.deepcopy(st.session_state.DEFAULT_JOBS_DB[selected_job])
        st.rerun()
    
    st.caption(f"当前加载: {st.session_state.current_job} 的配置")

    st.divider()

    # --- 规则展示 ---
    with st.expander("当前规则明细", expanded=True):
        rules_data = st.session_state.get("rules", {})
        if not rules_data:
            st.info("暂无规则")
        else:
            for category, rules_list in rules_data.items():
                st.markdown(f"**{category}**")
                if not rules_list:
                    st.caption("  (此类目下暂无规则)")
                for r in rules_list:
                    col1, col2 = st.columns([0.7, 0.3])
                    col1.caption(f"{r.get('keyword', '')} (+{r.get('weight', 0)})")
                    col2.caption(f"{r.get('desc', '')}")
                st.divider()

    st.divider()

    # --- 新增规则表单 ---
    st.subheader("➕ 新增规则")
    with st.form("add_rule_form"):
        new_cat = st.selectbox("所属分类", ["基本技能", "专业技能", "优选项"])
        new_kw = st.text_input("关键词")
        new_w = st.number_input("权重分", 1, 100, 5)
        new_desc = st.text_input("备注")
        
        if st.form_submit_button("添加", use_container_width=True):
            if new_kw:
                if new_cat not in st.session_state.rules:
                    st.session_state.rules[new_cat] = []
                st.session_state.rules[new_cat].append({"keyword": new_kw, "weight": new_w, "desc": new_desc})
                st.rerun()

    st.divider()

    # --- 重置与修改默认 (核心功能) ---
    st.subheader("🔄 重置与默认设置")
    
    col_reset, col_update = st.columns(2)
    
    with col_reset:
        if st.button("🗑️ 清除所有技能", use_container_width=True, type="secondary"):
            # 彻底清空当前岗位的规则
            st.session_state.rules = {"基本技能": [], "专业技能": [], "优选项": []}
            st.success("已清除所有技能！")
            st.rerun()
    
    with col_update:
        if st.button("💾 设为默认值", use_container_width=True, type="primary"):
            # 将当前修改过的规则保存为新的默认值
            st.session_state.DEFAULT_JOBS_DB[st.session_state.current_job] = copy.deepcopy(st.session_state.rules)
            st.success(f"已更新 {st.session_state.current_job} 的默认配置！")
            st.rerun()

# ---------- 5. 右侧：上传与筛选 ----------
with right:
    st.header(f"📂 {st.session_state.current_job} - 简历筛选")
    st.caption("支持包含子文件夹的 ZIP 包，自动识别 PDF/DOCX")
    
    up = st.file_uploader("选择 ZIP 文件", type=["zip"])

    # ---------- 工具函数 ----------
    def extract_text(path):
        try:
            if path.lower().endswith(".pdf"):
                with fitz.open(path) as doc:
                    return "\n".join([pg.get_text() for pg in doc]).lower()
            elif path.lower().endswith(".docx"):
                return "\n".join([p.text for p in Document(path).paragraphs]).lower()
        except Exception as e:
            st.warning(f"读取失败: {e}")
        return ""

    def score_it(text, rules_dict):
        total = 0
        detail = {}
        if not isinstance(rules_dict, dict):
            return {"total": 0, "detail": {}}
            
        for cat, rules in rules_dict.items():
            for r in rules:
                k = r.get("keyword", "")
                if not k: continue
                hit = bool(re.search(re.escape(k), text, re.IGNORECASE))
                s = r.get("weight", 0) if hit else 0
                detail[k] = s
                total += s
        return {"total": total, "detail": detail}

    # ---------- 筛选逻辑 ----------
    if st.button("🎯 开始智能筛选", use_container_width=True, type="primary"):
        if not up:
            st.warning("请上传 ZIP 文件")
        elif not any(st.session_state.rules.values()):
            st.warning("请先配置筛选规则")
        else:
            with tempfile.TemporaryDirectory() as td:
                zip_path = os.path.join(td, "bundle.zip")
                with open(zip_path, "wb") as f:
                    f.write(up.getvalue())

                extract_dir = os.path.join(td, "unzip")
                os.makedirs(extract_dir, exist_ok=True)
                
                with zipfile.ZipFile(zip_path, 'r') as zf:
                    for info in zf.infolist():
                        info.filename = info.filename.encode('cp437').decode('gbk')
                        zf.extract(info, extract_dir)

                targets = []
                for root, _, files in os.walk(extract_dir):
                    for fn in files:
                        if fn.lower().endswith((".pdf", ".docx")):
                            targets.append(os.path.join(root, fn))

                if not targets:
                    st.error("ZIP 内未找到 PDF/DOCX 文件")
                else:
                    rows = []
                    bar = st.progress(0.0, "解析中...")
                    for i, fp in enumerate(targets):
                        name = os.path.splitext(os.path.basename(fp))[0]
                        txt = extract_text(fp)
                        sc = score_it(txt, st.session_state.rules)
                        rec = {"简历文件": name, "总分": sc["total"]}
                        
                        for cat in st.session_state.rules.values():
                            for r in cat:
                                rec[r.get("keyword", "")] = sc["detail"].get(r.get("keyword", ""), 0)
                        
                        rows.append(rec)
                        bar.progress((i+1)/len(targets), name)
                    
                    bar.empty()
                    df = pd.DataFrame(rows).sort_values("总分", ascending=False).reset_index(drop=True)
                    st.success(f"✅ 完成，共筛选 {len(df)} 份简历")
                    
                    st.subheader("📊 Top 10 分数分布")
                    st.bar_chart(df.head(10).set_index("简历文件")["总分"])
                    
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    
                    csv = df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button("⬇ 下载完整结果 CSV", csv, "简历筛选结果.csv", "text/csv")
# ========== end ==========
