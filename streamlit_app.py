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
    .css-1aumxhk {
        background-color: #330000 !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 G-Change｜企業情報自動整形ツール")

st.markdown("企業リストの縦型データを1社1行に自動変換します。編集前ファイルをアップロードしてください。")

uploaded_file = st.file_uploader("📤 編集前のExcelファイルをアップロード", type=["xlsx"])

def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).strip()
    text = re.sub(r"[·⋅]", "", text)  # 中黒除去
    text = re.sub(r"\d+(\.\d+)?\([^)]+\)", "", text)  # 評価 5.0(4) など除去
    return text

def extract_info(group):
    texts = [clean_text(x) for x in group if pd.notna(x)]
    texts = [t for t in texts if t not in ["ウェブサイト", "ルート", "営業中", ""]]
    
    company = texts[0] if len(texts) > 0 else ""
    industry = texts[1] if len(texts) > 1 else ""
    address = texts[2] if len(texts) > 2 else ""
    
    # 電話番号を行全体から抽出
    phone = ""
    for t in texts:
        match = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", t)
        if match:
            phone = match.group()
            break
            
    return pd.Series([company, industry, address, phone])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    lines = df[0].dropna().tolist()

    # グループ化（空白行または「ルート」で区切る）
    groups = []
    current = []
    for line in lines:
        line = str(line).strip()
        if line in ["", "ルート"]:
            if current:
                groups.append(current)
                current = []
        else:
            current.append(line)
    if current:
        groups.append(current)

    result_df = pd.DataFrame([extract_info(group) for group in groups],
                             columns=["企業名", "業種", "住所", "電話番号"])

    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")

    st.dataframe(result_df, use_container_width=True)

    # ダウンロード用
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name="整形済みデータ")
    st.download_button("📥 Excelファイルをダウンロード", data=output.getvalue(),
                       file_name="整形済み_企業リスト.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")