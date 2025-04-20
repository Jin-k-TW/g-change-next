# 🚗 G-Change Next Ver3.9

import streamlit as st
import pandas as pd
import re
import io
import os
import shutil
from openpyxl import load_workbook

# ページ設定
st.set_page_config(page_title="G-Change Next", layout="wide")

# デザイン設定
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)

# タイトル
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver3.9）")

# --- NGリスト選択ブロック ---

nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
selected_nglist = st.selectbox("🛡️ 使用するNGリストを選択してください", nglist_options)

# --- 整形対象ファイルアップロードブロック ---

uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロード", type=["xlsx"])

# --- テキスト整形ルール ---

def normalize(text):
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    text = re.sub(r'[−–—―]', '-', text)
    return text

def extract_from_vertical_list(lines):
    extracted = []
    for i, line in enumerate(lines):
        if re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", str(line)):
            phone_line = normalize(str(line))
            phone_parts = phone_line.split("·")
            phone = phone_parts[-1].strip() if len(phone_parts) > 1 else phone_line.strip()

            address_line = normalize(str(lines[i-1])) if i-1 >= 0 else ""
            address_parts = address_line.split("·")
            address = address_parts[-1].strip() if len(address_parts) > 1 else address_line.strip()

            industry_line = normalize(str(lines[i-2])) if i-2 >= 0 else ""
            industry_parts = industry_line.split("·")
            industry = industry_parts[-1].strip() if len(industry_parts) > 1 else industry_line.strip()

            company = normalize(str(lines[i-3])) if i-3 >= 0 else ""

            extracted.append([company, industry, address, phone])
    return pd.DataFrame(extracted, columns=["企業名", "業種", "住所", "電話番号"])

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
        df = pd.read_excel(uploaded_file, sheet_name="入力マスター")
        if all(col in df.columns for col in ["企業様名称", "業種", "住所", "電話番号"]):
            result_df = df[["企業様名称", "業種", "住所", "電話番号"]].copy()
            result_df.columns = ["企業名", "業種", "住所", "電話番号"]
        else:
            st.error("⚠️ 入力マスターシートに必要な列（企業様名称、業種、住所、電話番号）がありません。")
            st.stop()
    else:
        df = pd.read_excel(uploaded_file, header=None)
        lines = df[0].dropna().tolist()
        result_df = extract_from_vertical_list(lines)

    result_df = clean_dataframe(result_df)

    company_removed = 0
    phone_removed = 0

    if selected_nglist != "なし":
        nglist_df = pd.read_excel(f"{selected_nglist}.xlsx")
        ng_companies = nglist_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        ng_phones = nglist_df.iloc[:, 1].dropna().astype(str).str.strip().tolist()

        before_company = len(result_df)
        result_df = result_df[~result_df["企業名"].apply(lambda x: any(ng_name in str(x) for ng_name in ng_companies))]
        after_company = len(result_df)
        company_removed = before_company - after_company

        before_phone = len(result_df)
        result_df = result_df[~result_df["電話番号"].astype(str).isin(ng_phones)]
        after_phone = len(result_df)
        phone_removed = before_phone - after_phone

    result_df = remove_phone_duplicates(result_df)
    result_df = remove_empty_rows(result_df)

    # --- 追加：電話番号で昇順ソート ---
    result_df = result_df.sort_values(by="電話番号", na_position='last').reset_index(drop=True)

    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")
    st.dataframe(result_df, use_container_width=True)

    if selected_nglist != "なし":
        st.info(f"🛡️ 【NGリスト削除件数】\n\n企業名による削除：{company_removed}件\n電話番号による削除：{phone_removed}件")

    template_file = "template.xlsx"
    output_file_name = f"{filename_no_ext}リスト.xlsx"
    shutil.copy(template_file, output_file_name)

    workbook = load_workbook(output_file_name)
    sheet = workbook["入力マスター"]

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row[1:]:
            cell.value = None

    for idx, row in result_df.iterrows():
        sheet.cell(row=idx+2, column=2, value=row["企業名"])
        sheet.cell(row=idx+2, column=3, value=row["業種"])
        sheet.cell(row=idx+2, column=4, value=row["住所"])
        sheet.cell(row=idx+2, column=5, value=row["電話番号"])

    workbook.save(output_file_name)

    with open(output_file_name, "rb") as f:
        st.download_button("📥 整形済みリストをダウンロード", data=f, file_name=output_file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
