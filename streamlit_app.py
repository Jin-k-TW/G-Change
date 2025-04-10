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
    h1 {
        color: #800000;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 G-Change｜企業情報自動整形ツール（Ver3）")

uploaded_file = st.file_uploader("📤 編集前のExcelファイルをアップロード", type=["xlsx"])

# ヘルパー関数群
def is_company_name(line):
    return "·" not in line and "⋅" not in line and "：" not in line and not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

def normalize(text):
    return re.sub(r'[−–—―]', '-', str(text)).strip()

def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""
    review_keywords = ["できる", "優しい", "楽しい", "助かる", "人柄", "感じ", "スタッフ", "雰囲気"]
    ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

    for line in lines[1:]:
        line = normalize(str(line))
        if any(kw in line for kw in ignore_keywords):
            continue
        if any(kw in line for kw in review_keywords):
            continue
        if "·" in line or "⋅" in line:
            parts = re.split(r"[·⋅]", line)
            industry = parts[-1].strip()
        elif re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line) or re.search(r"\d{10,11}", line.replace("-", "").replace(" ", "")):
            phone_match = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)
            if not phone_match:
                digits = re.sub(r"[^\d]", "", line)
                if len(digits) >= 10:
                    phone = f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
            else:
                phone = phone_match.group()
        elif not address and not industry and not re.search(r"[ぁ-んァ-ン]", line):
            address = line  # 住所候補

    return pd.Series([company, industry, address, phone])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)
    lines = df[0].dropna().tolist()

    # 企業ごとのグループ化
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

    # ダウンロード
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name="整形済みデータ")
    st.download_button("📥 Excelファイルをダウンロード", data=output.getvalue(),
                       file_name="整形済み_企業リスト.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")