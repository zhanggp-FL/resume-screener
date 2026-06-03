import streamlit as st
import os
import fitz
from docx import Document
import pandas as pd

st.set_page_config(page_title="简历筛选器", layout="wide")
st.title("📄 光学工程师简历筛选系统")

# 规则（和之前一样）
RULES = {
    "学历": {"硕士": 10, "博士": 12, "研究生": 8},
    "理论知识": {"几何光学": 6, "物理光学": 6, "激光原理": 6},
    "光学工具": {"zemax": 5, "matlab": 5, "virtuallab": 5},
    "编程能力": {"python": 5, "c#": 4, ".net": 3, "宏": 3, "二次开发": 4},
    "工艺经验": {"光学加工": 4, "装调": 4, "工艺": 3},
    "软素质": {"团队合作": 2, "创新": 2, "解决问题": 2},
    "优先项_DOE": {"doe": 6, "衍射光学": 6},
    "语言能力": {"英语": 3, "english": 3}
}


def extract_text(path):
    text = ""
    if path.endswith(".pdf"):
        with fitz.open(path) as doc:
            for page in doc: text += page.get_text()
    elif path.endswith(".docx"):
        doc = Document(path)
        text = "\n".join([p.text for p in doc.paragraphs])
    return text.lower()


def calc_score(text):
    detail = {}
    total = 0
    for cat, kws in RULES.items():
        score = sum(w for k, w in kws.items() if k in text)
        detail[cat] = score
        total += score
    return total, detail


# 上传文件
uploaded_files = st.file_uploader("上传简历（支持多选）", accept_multiple_files=True, type=['pdf', 'docx'])

if uploaded_files and st.button("开始筛选"):
    results = []
    for uf in uploaded_files:
        # 保存到临时文件
        temp_path = f"temp_{uf.name}"
        with open(temp_path, "wb") as f:
            f.write(uf.getbuffer())

        text = extract_text(temp_path)
        total, detail = calc_score(text)
        results.append({"姓名": uf.name, "总分": total, **detail})
        os.remove(temp_path)

    df = pd.DataFrame(results)
    st.success(f"筛选完成，共处理 {len(df)} 份简历")
    st.dataframe(df.sort_values(by="总分", ascending=False))


    # 下载按钮
    @st.cache_data
    def convert_df(df):
        return df.to_csv(index=False).encode('utf-8-sig')


    csv = convert_df(df)
    st.download_button("下载 CSV", csv, "结果.csv", "text/csv")