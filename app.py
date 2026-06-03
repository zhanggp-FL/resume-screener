# ========== app.py (修复乱码版) ==========
import streamlit as st
import re
import pandas as pd
import tempfile
import os
import zipfile

import fitz
from docx import Document

st.set_page_config(page_title="简历筛选系统", layout="wide")
st.title("简历筛选系统")

# ---------- 规则初始化 ----------
DEFAULT_RULES = [
    {"keyword": "Zemax", "weight": 8, "desc": "熟练使用 Zemax / CodeV"},
    {"keyword": "Python", "weight": 5, "desc": "具备 Python 编程能力"},
    {"keyword": "光学设计", "weight": 10, "desc": "几何光学/成像光学/照明光学设计经验"},
]

if "rules" not in st.session_state:
    st.session_state.rules = list(DEFAULT_RULES)

# ---------- 左侧：规则管理 ----------
left, right = st.columns([1, 2])

with left:
    st.header("⚙️ 筛选规则")
    with st.expander("当前规则", expanded=True):
        rows = []
        for i, r in enumerate(st.session_state.rules):
            rows.append({"序号": i+1, "关键词": r["keyword"], "权重分": r["weight"], "备注": r["desc"]})
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.info("暂无规则")

    st.subheader("新增 / 覆盖更新规则")
    kw = st.text_input("关键词", key="kw")
    w = st.number_input("权重分", 1, 100, 5, key="w")
    dc = st.text_input("备注", key="dc")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("➕ 保存规则", use_container_width=True, type="primary"):
            if not kw.strip():
                st.warning("关键词不能为空")
            else:
                k = kw.strip()
                st.session_state.rules = [r for r in st.session_state.rules if r["keyword"] != k]
                st.session_state.rules.append({"keyword": k, "weight": int(w), "desc": dc})
                st.rerun()
    with c2:
        if st.button("🗑️ 清空规则", use_container_width=True):
            st.session_state.rules = []
            st.rerun()

# ---------- 右侧：上传与筛选 ----------
with right:
    st.header("📂 上传简历（ZIP 压缩包）")
    up = st.file_uploader("选择 ZIP 文件", type=["zip"])

    # ---------- 修复点：文件名解码函数 ----------
    def decode_filename(filename_bytes):
        """
        尝试用多种编码解码文件名，解决中文乱码问题
        """
        encodings_to_try = ['utf-8', 'gbk', 'cp936', 'big5']
        for enc in encodings_to_try:
            try:
                return filename_bytes.decode(enc)
            except UnicodeDecodeError:
                continue
        # 如果都失败了，强行替换乱码字符
        return filename_bytes.decode('utf-8', errors='replace')

    # ---------- 文本提取 ----------
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

    # ---------- 打分 ----------
    def score_it(text, rules):
        detail = {}
        total = 0
        for r in rules:
            k = r["keyword"]
            hit = bool(re.search(re.escape(k), text, re.IGNORECASE))
            s = r["weight"] if hit else 0
            detail[k] = s
            total += s
        return {"total": total, "detail": detail}

    # ---------- 开始筛选 ----------
    if st.button("🎯 开始筛选", use_container_width=True, type="primary"):
        if not up:
            st.warning("请上传 ZIP")
        elif not st.session_state.rules:
            st.warning("请添加规则")
        else:
            with tempfile.TemporaryDirectory() as td:
                zip_path = os.path.join(td, "bundle.zip")
                with open(zip_path, "wb") as f:
                    f.write(up.getvalue())

                extract_dir = os.path.join(td, "unzip")
                os.makedirs(extract_dir, exist_ok=True)
                
                # ---------- 修复点：使用自定义解压逻辑 ----------
                with zipfile.ZipFile(zip_path, 'r') as zf:
                    for info in zf.infolist():
                        # 这里是关键：解码文件名
                        real_filename = decode_filename(info.filename.encode('cp437'))
                        info.filename = real_filename
                        zf.extract(info, extract_dir)

                targets = []
                for root, _, files in os.walk(extract_dir):
                    for fn in files:
                        if fn.lower().endswith((".pdf", ".docx")):
                            targets.append(os.path.join(root, fn))

                if not targets:
                    st.error("ZIP 内未找到 PDF/DOCX")
                else:
                    rows = []
                    bar = st.progress(0.0, "解析中...")
                    for i, fp in enumerate(targets):
                        name = os.path.splitext(os.path.basename(fp))[0]
                        txt = extract_text(fp)
                        sc = score_it(txt, st.session_state.rules)
                        rec = {"简历文件": name, "总分": sc["total"]}
                        for r in st.session_state.rules:
                            rec[r["keyword"]] = sc["detail"].get(r["keyword"], 0)
                        rows.append(rec)
                        bar.progress((i+1)/len(targets), name)
                    
                    bar.empty()
                    df = pd.DataFrame(rows).sort_values("总分", ascending=False).reset_index(drop=True)
                    st.success(f"完成，共 {len(df)} 份")
                    st.dataframe(df, use_container_width=True, hide_index=True)

                    csv = df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button("⬇ 下载结果", csv, "简历筛选结果.csv", "text/csv")
# ========== end ==========
