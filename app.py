import streamlit as st
import pandas as pd
import tempfile
import os
import zipfile
import copy
import io
import base64
from pathlib import Path

import fitz
from docx import Document

# 支持 .7z 格式
try:
    import py7zr
except ImportError:
    st.error("请安装 py7zr 以支持 .7z 文件: pip install py7zr")
    st.stop()

# ---------- 1. 初始默认规则 ----------
INITIAL_DEFAULTS = {
    "光学设计工程师": {
        "基本技能": [{"keyword": "光学设计", "weight": 10, "desc": "几何/成像"},
                     {"keyword": "Zemax", "weight": 8, "desc": "软件"}],
        "专业技能": [{"keyword": "DOE", "weight": 6, "desc": "衍射"},
                     {"keyword": "杂散光", "weight": 5, "desc": "分析"}],
        "优选项": [{"keyword": "Python", "weight": 4, "desc": "自动化"},
                   {"keyword": "Matlab", "weight": 3, "desc": "算法"}]
    },
    "机械结构设计工程师": {
        "基本技能": [{"keyword": "SolidWorks", "weight": 10, "desc": "建模"},
                     {"keyword": "工程图", "weight": 8, "desc": "出图"}],
        "专业技能": [{"keyword": "公差分析", "weight": 7, "desc": "GD&T"},
                     {"keyword": "热设计", "weight": 5, "desc": "散热"}],
        "优选项": [{"keyword": "ANSYS", "weight": 4, "desc": "有限元"},
                   {"keyword": "项目管理", "weight": 3, "desc": "协作"}]
    },
    "CAE工程师": {
        "基本技能": [{"keyword": "Hypermesh", "weight": 10, "desc": "前处理"},
                     {"keyword": "Abaqus", "weight": 8, "desc": "求解"}],
        "专业技能": [{"keyword": "模态分析", "weight": 7, "desc": "动力学"},
                     {"keyword": "疲劳分析", "weight": 6, "desc": "寿命"}],
        "优选项": [{"keyword": "Python", "weight": 5, "desc": "二次开发"},
                   {"keyword": "二次开发", "weight": 4, "desc": "脚本"}]
    }
}

# ---------- 2. Session State 初始化 ----------
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

    job_options = list(st.session_state.DEFAULT_JOBS_DB.keys())
    current_idx = job_options.index(st.session_state.current_job)

    selected_job = st.selectbox("选择招聘岗位", options=job_options, index=current_idx)

    if selected_job != st.session_state.current_job:
        st.session_state.current_job = selected_job
        st.session_state.rules = copy.deepcopy(st.session_state.DEFAULT_JOBS_DB[selected_job])
        st.rerun()

    st.caption(f"当前加载: {st.session_state.current_job} 的配置")

    st.divider()

    # 显示规则明细（带删除按钮）
    with st.expander("当前规则明细", expanded=True):
        rules_data = st.session_state.get("rules", {})
        if not rules_data or not any(rules_data.values()):
            st.info("暂无规则")
        else:
            for category, rules_list in rules_data.items():
                st.markdown(f"**{category}**")
                if not rules_list:
                    st.caption("  (此类目下暂无规则)")
                else:
                    for idx, r in enumerate(rules_list):
                        col1, col2, col3 = st.columns([0.5, 0.3, 0.2])
                        col1.caption(f"{r.get('keyword', '')} (+{r.get('weight', 0)})")
                        col2.caption(r.get('desc', ''))
                        if col3.button("❌", key=f"del_{category}_{idx}"):
                            st.session_state.rules[category].pop(idx)
                            st.session_state.DEFAULT_JOBS_DB[st.session_state.current_job] = copy.deepcopy(st.session_state.rules)
                            st.rerun()
                st.divider()

    st.divider()

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
                st.session_state.rules[new_cat].append({
                    "keyword": new_kw,
                    "weight": new_w,
                    "desc": new_desc
                })
                st.session_state.DEFAULT_JOBS_DB[st.session_state.current_job] = copy.deepcopy(st.session_state.rules)
                st.rerun()

    st.divider()

    st.subheader("🔄 重置与默认设置")
    col_reset, col_update = st.columns(2)

    with col_reset:
        if st.button("🗑️ 清除所有技能", use_container_width=True, type="secondary"):
            st.session_state.rules = {"基本技能": [], "专业技能": [], "优选项": []}
            st.session_state.DEFAULT_JOBS_DB[st.session_state.current_job] = copy.deepcopy(st.session_state.rules)
            st.success("已清除所有技能！")
            st.rerun()

    with col_update:
        if st.button("💾 设为默认值", use_container_width=True, type="primary"):
            st.session_state.DEFAULT_JOBS_DB[st.session_state.current_job] = copy.deepcopy(st.session_state.rules)
            st.success(f"已更新 {st.session_state.current_job} 的默认配置！")
            st.rerun()

# ---------- 5. 右侧：上传与筛选 ----------
with right:
    st.header(f"📂 {st.session_state.current_job} - 简历筛选")
    st.caption("支持 ZIP 或 7Z 格式（自动解压，包含子文件夹）")

    up = st.file_uploader("选择压缩文件", type=["zip", "7z"])

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
        for cat, rules in rules_dict.items():
            for r in rules:
                kw = r.get("keyword", "")
                if not kw:
                    continue
                hit = kw.lower() in text
                s = r.get("weight", 0) if hit else 0
                detail[kw] = detail.get(kw, 0) + s
                total += s
        return {"total": total, "detail": detail}

    def build_all_keywords(rules_dict):
        keywords = set()
        for cat, rules in rules_dict.items():
            for r in rules:
                kw = r.get("keyword", "")
                if kw:
                    keywords.add(kw)
        return sorted(keywords)

    def make_result_zip(file_paths, output_zip_path):
        with zipfile.ZipFile(output_zip_path, 'w') as zf:
            for abs_path, arcname in file_paths:
                zf.write(abs_path, arcname)

    # ---------- 筛选逻辑 ----------
    if st.button("🎯 开始智能筛选", use_container_width=True, type="primary"):
        if not up:
            st.warning("请上传 ZIP 或 7Z 文件")
        elif not any(st.session_state.rules.values()):
            st.warning("请先配置筛选规则")
        else:
            with tempfile.TemporaryDirectory() as td:
                # 保存上传的压缩包
                file_ext = os.path.splitext(up.name)[1].lower()
                compressed_path = os.path.join(td, f"bundle{file_ext}")
                with open(compressed_path, "wb") as f:
                    f.write(up.getvalue())

                extract_dir = os.path.join(td, "unzip")
                os.makedirs(extract_dir, exist_ok=True)

                # 解压
                try:
                    if file_ext == '.zip':
                        with zipfile.ZipFile(compressed_path, 'r') as zf:
                            for info in zf.infolist():
                                try:
                                    info.filename = info.filename.encode('cp437').decode('gbk')
                                except:
                                    pass
                                zf.extract(info, extract_dir)
                    elif file_ext == '.7z':
                        with py7zr.SevenZipFile(compressed_path, mode='r') as zf:
                            zf.extractall(extract_dir)
                    else:
                        st.error("不支持的文件格式")
                        st.stop()
                except Exception as e:
                    st.error(f"解压失败: {e}")
                    st.stop()

                # 收集所有 PDF/DOCX
                targets = []  # (相对路径, 绝对路径)
                for root, _, files in os.walk(extract_dir):
                    for fn in files:
                        if fn.lower().endswith((".pdf", ".docx")):
                            abs_path = os.path.join(root, fn)
                            rel_path = os.path.relpath(abs_path, extract_dir)
                            targets.append((rel_path, abs_path))

                if not targets:
                    st.error("压缩包内未找到 PDF/DOCX 文件")
                else:
                    all_keywords = build_all_keywords(st.session_state.rules)
                    rows = []
                    bar = st.progress(0.0, "解析中...")
                    qualified_for_zip = []  # 总分>0的简历，用于打包下载
                    file_links = []          # 存储每个文件的 base64 下载链接（与 targets 顺序对应）

                    for i, (rel_path, abs_path) in enumerate(targets):
                        name = os.path.splitext(rel_path)[0].replace(os.sep, '/')
                        txt = extract_text(abs_path)
                        sc = score_it(txt, st.session_state.rules)

                        rec = {"简历文件": name, "总分": sc["total"]}
                        for kw in all_keywords:
                            rec[kw] = sc["detail"].get(kw, 0)
                        rows.append(rec)

                        if sc["total"] > 0:
                            qualified_for_zip.append((abs_path, rel_path))

                        # 生成 HTML 下载链接（base64）
                        with open(abs_path, "rb") as f:
                            file_bytes = f.read()
                        b64 = base64.b64encode(file_bytes).decode()
                        mime = "application/pdf" if abs_path.lower().endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        href = f'<a href="data:{mime};base64,{b64}" download="{os.path.basename(abs_path)}" target="_blank">📥 下载</a>'
                        file_links.append(href)

                        bar.progress((i + 1) / len(targets), name)

                    bar.empty()
                    df = pd.DataFrame(rows).sort_values("总分", ascending=False).reset_index(drop=True)
                    df.insert(0, "排名", range(1, len(df) + 1))

                    st.success(f"✅ 完成，共筛选 {len(df)} 份简历")

                    # 显示 Top 10 分数分布
                    st.subheader("📊 Top 10 分数分布")
                    st.bar_chart(df.head(10).set_index("简历文件")["总分"])

                    # 显示带下载链接的 HTML 表格
                    st.subheader("📋 筛选结果（可点击下载简历）")
                    # 重新排序链接以匹配排序后的 df
                    sorted_indices = df.index
                    sorted_links = [file_links[idx] for idx in sorted_indices]
                    html_df = df.copy()
                    html_df["下载简历"] = sorted_links
                    display_cols = ["排名", "简历文件", "总分"] + all_keywords + ["下载简历"]
                    st.markdown(html_df[display_cols].to_html(escape=False, index=False), unsafe_allow_html=True)

                    # 下载 CSV（不含链接）
                    csv_data = df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button("⬇ 下载结果 CSV", data=csv_data, file_name="简历筛选结果.csv", mime="text/csv")

                    # 下载 Excel（不含链接，但包含所有数据 + 排名）
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="简历评分")
                    excel_data = output.getvalue()
                    st.download_button("⬇ 下载结果 Excel", data=excel_data, file_name="简历筛选结果.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                    # 打包下载筛选后的简历（总分>0）
                    if qualified_for_zip:
                        with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp_zip:
                            zip_path_tmp = tmp_zip.name
                        make_result_zip(qualified_for_zip, zip_path_tmp)
                        with open(zip_path_tmp, "rb") as f:
                            zip_bytes = f.read()
                        st.download_button("📦 打包下载筛选后的简历 (ZIP)", data=zip_bytes,
                                           file_name=f"{st.session_state.current_job}_筛选简历.zip",
                                           mime="application/zip")
                        os.unlink(zip_path_tmp)
                    else:
                        st.info("没有总分大于0的简历，无法打包下载。")
