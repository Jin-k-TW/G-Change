import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="G-Change", layout="wide")

st.markdown("""
    <style>
    body, .stApp {
        background-color: #f9f9f9;
        font-family: 'Helvetica Neue', sans-serif;
    }
    .main {
        color: #330000;
    }
    h1 {
        color: #800000;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 G-Change｜企業情報自動整形ツール")

st.markdown("企業リストの縦型データを1社1行に自動変換します。編集前ファイルをアップロードしてください。")

uploaded_file = st.file_uploader("📤 編集前のExcelファイルをアップロード", type=["xlsx"])

def is_company_name(line):
    return "·" not in line and "⋅" not in line and "：" not in line and not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

def extract_info(lines):
    company = lines[0].strip() if lines else ""
    industry, address, phone = "", "", ""

    for line in lines[1:]:
        line = str(line).strip()
        if "·" in line or "⋅" in line:
            parts = re.split(r"[·⋅]", line)
            for part in parts:
                if "メーカー" in part or "工業" in part or "店" in part or "業" in part:
                    industry = part.strip()
        elif re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line):
            phone_match = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)
            if phone_match:
                phone = phone_match.group()
        elif any(keyword in line for keyword in ["町", "丁目", "番", "−", "-"]):
            address = line.strip()

    return pd.Series([company, industry, address, phone])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    lines = df[0].dropna().tolist()

    groups = []
    current = []
    for line in lines:
        line = str(line).strip()
        if is_company_name(line):
            if current:
                groups.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        groups.append(current)

    result_df = pd.DataFrame([extract_info(group) for group in groups],
                             columns=["企業名", "業種", "住所", "電話番号"])

    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")
    st.dataframe(result_df, use_container_width=True)

    # Excelダウンロード
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name="整形済みデータ")
    st.download_button("📥 Excelファイルをダウンロード", data=output.getvalue(),
                       file_name="整形済み_企業リスト.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
