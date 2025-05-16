import streamlit as st
import pandas as pd
import re
import io
import os
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook

# ページ設定
st.set_page_config(page_title="G-Change Next", layout="wide")

# デザイン設定
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver4.2 安定版）")

# --- NGリスト選択 ---
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
selected_nglist = st.selectbox("🛡️ 使用するNGリストを選択してください", nglist_options)

# --- ファイルアップロード ---
uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロード", type=["xlsx"])

# --- ユーティリティ関数群 ---
def normalize(text):
    if text is None or pd.isna(text):
        return ""
    text = str(text).strip().replace("\u3000", " ").replace("\xa0", " ")
    text = re.sub(r'[−–—―]', '-', text)
    return text

def is_phone(line):
    return re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

def extract_company_groups(lines):
    results = []
    buffer = []
    for line in lines:
        line = normalize(str(line))
        if not line or line in ["ルート", "ユーザーサイト"]:
            continue
        buffer.append(line)
        if is_phone(line):
            phone_match = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)
            phone = phone_match.group() if phone_match else ""

            address = ""
            industry = ""
            company = ""

            for back_line in reversed(buffer[:-1][-6:]):
                if not address and any(x in back_line for x in ["丁目", "区", "市", "番地", "-", "−"]):
                    address = back_line
                elif not industry and any(x in back_line for x in ["プラスチック", "製造", "加工", "業", "サービス"]):
                    industry = back_line
                elif not company:
                    company = back_line

            results.append([company, industry, address, phone])
            buffer.clear()
    return pd.DataFrame(results, columns=["企業名", "業種", "住所", "電話番号"])

def clean_dataframe(df):
    return df.applymap(lambda x: str(x).strip() if pd.notnull(x) else x)

def remove_phone_duplicates(df):
    seen_phones = set()
    cleaned_rows = []
    for _, row in df.iterrows():
        phone = str(row["電話番号"]).strip()
        if phone == "" or phone not in seen_phones:
            cleaned_rows.append(row)
            if phone != "":
                seen_phones.add(phone)
    return pd.DataFrame(cleaned_rows)

def remove_empty_rows(df):
    return df[~((df["企業名"] == "") & (df["業種"] == "") & (df["住所"] == "") & (df["電話番号"] == ""))]

# --- 実行メインブロック ---
if uploaded_file:
    filename_no_ext = os.path.splitext(uploaded_file.name)[0]
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names

    if "入力マスター" in sheet_names:
        df_raw = pd.read_excel(uploaded_file, sheet_name="入力マスター", header=None)
        result_df = pd.DataFrame({
            "企業名": df_raw.iloc[:, 1].astype(str).apply(normalize),
            "業種": df_raw.iloc[:, 2].astype(str).apply(normalize),
            "住所": df_raw.iloc[:, 3].astype(str).apply(normalize),
            "電話番号": df_raw.iloc[:, 4].astype(str).apply(normalize)
        })
    else:
        df = pd.read_excel(uploaded_file, header=None)
        lines = df[0].dropna().tolist()
        result_df = extract_company_groups(lines)

    result_df = clean_dataframe(result_df)

    company_removed = 0
    phone_removed = 0
    if selected_nglist != "なし":
        ng_path = f"{selected_nglist}.xlsx"
        if not os.path.exists(ng_path):
            st.error(f"❌ 選択されたNGリストファイルが見つかりません：{ng_path}")
            st.stop()
        ng_df = pd.read_excel(ng_path)
        if ng_df.shape[1] < 2:
            st.error("❌ NGリストは2列以上必要です（企業名、電話番号）")
            st.stop()
        ng_companies = ng_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        ng_phones = ng_df.iloc[:, 1].dropna().astype(str).str.strip().tolist()

        before_company = len(result_df)
