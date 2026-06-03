# ========== app.py ==========
import streamlit as st
import re
import pandas as pd
import tempfile
import os
import zipfile

# 文本提取用到
import fitz  # PyMuPDF
from docx import Document

# ---------- 页面配置 ----------
st.set_page_config(page_title="简历筛选系统", layout="wide")
st.title("简历筛选系统")

# ---------- 规则初始化 ----------
DEFAULT_RULES = [
    {"keyword": "Zemax",    "weight": 8, "desc": "熟练使用 Zemax / CodeV"},
    {"keyword": "Python",   "weight": 5, "desc": "具备 Python 编程能力"},
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
            rows.append({
                "序号": i + 1,
                "关键词": r["keyword"],
                "权重分": r["weight"],
                "备注": r["desc"],
            })
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.info("暂无规则，请在下方添加")

    st.subheader("新增 / 覆盖更新规则")

    kw = st.text_input("关键词（如 Zemax / DOE / 激光原理）", key="kw")
    w  = st.number_input("权重分（命中该关键词加几分）", min_value=1, max_value=100, value=5, step=1, key="w")
    dc = st.text_input("备注（可选）", key="dc")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("➕ 保存规则（同名自动覆盖）", use_container_width=True, type="primary"):
            if not kw.strip():
                st.warning("关键词不能为空")
            else:
                k = kw.strip()
                # 若同名已存在：先移除旧条目（覆盖语义）
                st.session_state.rules = [r for r in st.session_state.rules if r["keyword"] != k]
                st.session_state.rules.append({"keyword": k, "weight": int(w), "desc": dc})
                st.success("已保存：" + k)
                st.rerun()
    with c2:
        if st.button("🗑️ 清空所有规则", use_container_width=True):
            st.session_state.rules = []
            st.rerun()

# ---------- 右侧：上传 ZIP + 筛选 ----------
with right:
    st.header("📂 上传简历（ZIP 压缩包）")
    st.caption("把你的“resumes 总文件夹”右键 → 压缩成 ZIP，这里上传即可（子文件夹也会被扫描）")

    up = st.file_uploader(
        "选择 ZIP 文件",
        type=["zip"],
    )

    # ---- 工具函数：文本提取 ----
    def _read_pdf(path: str) -> str:
        out = []
        with fitz.open(path) as doc:
            for pg in doc:
                out.append(pg.get_text())
        return "\n".join(out)

    def _read_docx(path: str) -> str:
        parms = []
        d = Document(path)
        for p in d.paragraphs:
            if p.text:
                parms.append(p.text)
        return "\n".join(parms)

    def extract_text_from_path(path: str) -> str:
        t = ""
        try:
            if path.lower().endswith(".pdf"):
                t = _read_pdf(path)
            elif path.lower().endswith(".docx"):
                t = _read_docx(path)
        except Exception as exc:
            # 这里一定用双引号闭合，避免 unterminated string
            st.warning("读取失败: " + os.path.basename(path) + " (" + str(exc) + ")")
        return t.lower()

    # ---- 关键词/规则打分 ----
    def score_it(text: str, rules: list) -> dict:
        detail = {}
        total = 0
        for r in rules:
            k = (r.get("keyword") or "").strip()
            if not k:
                detail["(空)"] = 0
                continue
            hit = 1 if re.search(re.escape(k), text, re.IGNORECASE) else 0
            s = r["weight"] if hit else 0
            detail[k] = s
            total += s
        return {"total": total, "detail": detail}

    # ---- 开始筛选 ----
    if st.button("🎯 开始筛选", use_container_width=True, type="primary"):
        if not up:
            st.warning("请先上传 ZIP 文件")
        elif not st.session_state.rules:
            st.warning("请先添加至少一条筛选规则")
        else:
            with tempfile.TemporaryDirectory() as td:
                zip_path = os.path.join(td, "bundle.zip")
                with open(zip_path, "wb") as fout:
                    fout.write(up.getvalue())

                extract_dir = os.path.join(td, "unzip")
                os.makedirs(extract_dir, exist_ok=True)
                with zipfile.ZipFile(zip_path, "r") as zf:
                    zf.extractall(extract_dir)

                targets = []
                for root, _, files in os.walk(extract_dir):
                    for fn in files:
                        if fn.lower().endswith((".pdf", ".docx")):
                            targets.append(os.path.join(root, fn))

                if not targets:
                    st.error("ZIP 内未找到 PDF / DOCX 文件")
                else:
                    rows_out = []
                    bar = st.progress(0.0, text="解析中...")
                    for idx, fp in enumerate(targets):
                        txt = extract_text_from_path(fp)
                        sc = score_it(txt, st.session_state.rules)
                        name = os.path.splitext(os.path.basename(fp))[0]
                        rec = {"简历文件": name, "总分": sc["total"]}
                        for r in st.session_state.rules:
                            k = (r.get("keyword") or "").strip()
                            rec[k] = sc["detail"].get(k, 0)
                        rows_out.append(rec)
                        bar.progress((idx + 1) / max(len(targets), 1), text=name)

                    bar.empty()
                    df = pd.DataFrame(rows_out).sort_values("总分", ascending=False).reset_index(drop=True)
                    st.success("完成，共 {} 份".format(len(df)))
                    st.dataframe(df, use_container_width=True, hide_index=True)

                    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        "⬇ 下载结果 CSV（Excel 可直接打开）",
                        data=csv_bytes,
                        file_name="简历筛选结果.csv",
                        mime="text/csv",
                    )
# ========== end app.py ==========
