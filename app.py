# ========== app.py (最终稳定版) ==========
import streamlit as st
import re
import pandas as pd
import tempfile
import os
import zipfile

import fitz
from docx import Document

# ---------- 1. 多岗位规则库 (核心数据) ----------
JOBS_DB = {
    "光学设计工程师": {
        "基本技能": [
            {"keyword": "光学设计", "weight": 10, "desc": "几何/成像/照明设计"},
            {"keyword": "Zemax", "weight": 8, "desc": "光学仿真软件"},
        ],
        "专业技能": [
            {"keyword": "DOE", "weight": 6, "desc": "衍射光学元件"},
            {"keyword": "杂散光", "weight": 5, "desc": "Stray Light Analysis"},
        ],
        "优选项": [
            {"keyword": "Python", "weight": 4, "desc": "二次开发或自动化"},
            {"keyword": "Matlab", "weight": 3, "desc": "算法验证"},
        ]
    },
    "机械结构设计工程师": {
        "基本技能": [
            {"keyword": "SolidWorks", "weight": 10, "desc": "3D建模软件"},
            {"keyword": "工程图", "weight": 8, "desc": "出图能力"},
        ],
        "专业技能": [
            {"keyword": "公差分析", "weight": 7, "desc": "GD&T"},
            {"keyword": "热设计", "weight": 5, "desc": "散热结构"},
        ],
        "优选项": [
            {"keyword": "ANSYS", "weight": 4, "desc": "有限元分析"},
            {"keyword": "项目管理", "weight": 3, "desc": "跨部门协作"},
        ]
    },
    "CAE工程师": {
        "基本技能": [
            {"keyword": "Hypermesh", "weight": 10, "desc": "前处理网格划分"},
            {"keyword": "Abaqus", "weight": 8, "desc": "非线性求解"},
        ],
        "专业技能": [
            {"keyword": "模态分析", "weight": 7, "desc": "动力学仿真"},
            {"keyword": "疲劳分析", "weight": 6, "desc": "寿命预测"},
        ],
        "优选项": [
            {"keyword": "Python", "weight": 5, "desc": "二次开发"},
            {"keyword": "二次开发", "weight": 4, "desc": "脚本编写"},
        ]
    }
}

# ---------- 2. Session State 初始化 (防崩溃核心) ----------
if "current_job" not in st.session_state:
    st.session_state.current_job = list(JOBS_DB.keys())[0]  # 默认选第一个

if "rules" not in st.session_state:
    # 确保 rules 初始化为一个有效的字典，而不是空列表
    st.session_state.rules = JOBS_DB[st.session_state.current_job]

# ---------- 3. 页面配置 ----------
st.set_page_config(page_title="多岗位简历筛选系统", layout="wide")
st.title("📄 多岗位简历智能筛选系统")

left, right = st.columns([1, 2])

# ---------- 4. 左侧：岗位切换与规则管理 ----------
with left:
    st.header("⚙️ 岗位与规则配置")
    
    # --- 岗位选择器 (修复后的逻辑) ---
    job_options = list(JOBS_DB.keys())
    current_idx = job_options.index(st.session_state.get("current_job", job_options[0]))
    
    selected_job = st.selectbox(
        "选择招聘岗位",
        options=job_options,
        index=current_idx
    )
    
    # 只有当下拉框的值和 session_state 不一致时才更新，防止无限 rerun
    if selected_job != st.session_state.current_job:
        st.session_state.current_job = selected_job
        st.session_state.rules = JOBS_DB[selected_job]
        st.rerun()
    
    st.caption(f"当前加载: {st.session_state.current_job} 的配置")

    st.divider()

    # --- 规则展示 (安全读取) ---
    with st.expander("当前规则明细", expanded=True):
        # 使用 .get() 方法安全获取 rules，如果不存在则返回空字典
        rules_data = st.session_state.get("rules", {})
        if not rules_data:
            st.info("暂无规则，请选择岗位或添加规则。")
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
                # 确保 rules 是字典且包含该分类
                if new_cat not in st.session_state.rules:
                    st.session_state.rules[new_cat] = []
                
                st.session_state.rules[new_cat].append({
                    "keyword": new_kw,
                    "weight": new_w,
                    "desc": new_desc
                })
                st.rerun()

# ---------- 5. 右侧：上传与筛选 ----------
with right:
    st.header(f"📂 {st.session_state.get('current_job', '未知岗位')} - 简历筛选")
    st.caption("支持包含子文件夹的 ZIP 包，自动识别 PDF/DOCX")
    
    up = st.file_uploader("选择 ZIP 文件", type=["zip"])

    # ---------- 工具函数 ----------
    def extract_text(path):
        try:
            if path.lower().endswith(".pdf"):
                with fitz.open(path) as doc:
                    return "\n".join([pg.get_text() for pg in doc]).lower()
            elif path.lower().endswith(".docx"):
                doc = Document(path)
                return "\n".join([p.text for p in doc.paragraphs]).lower()
        except Exception as e:
            st.warning(f"读取失败: {os.path.basename(path)} ({e})")
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
        elif not st.session_state.get("rules"):
            st.warning("请先配置筛选规则")
        else:
            with tempfile.TemporaryDirectory() as td:
                zip_path = os.path.join(td, "bundle.zip")
                with open(zip_path, "wb") as f:
                    f.write(up.getvalue())

                extract_dir = os.path.join(td, "unzip")
                os.makedirs(extract_dir, exist_ok=True)
                
                # 修复中文乱码
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
                        
                        # 扁平化处理
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
