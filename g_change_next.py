import streamlit as st
import pandas as pd
import re
import io
import os
from openpyxl import load_workbook

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
        result_df = result_df[~result_df["企業名"].apply(lambda x: any(ng in str(x) for ng in ng_companies))]
        company_removed = before_company - len(result_df)

        before_phone = len(result_df)
        result_df = result_df[~result_df["電話番号"].astype(str).isin(ng_phones)]
        phone_removed = before_phone - len(result_df)

    result_df = remove_phone_duplicates(result_df)
    result_df = remove_empty_rows(result_df)
    result_df = result_df.sort_values(by="電話番号", na_position='last').reset_index(drop=True)

    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")
    st.dataframe(result_df, use_container_width=True)

    if selected_nglist != "なし":
        st.info(f"🛡️ 【NGリスト削除件数】\n\n企業名による削除：{company_removed}件\n電話番号による削除：{phone_removed}件")

    template_file = "template.xlsx"
    if not os.path.exists(template_file):
        st.error("❌ template.xlsx が存在しません")
        st.stop()

    workbook = load_workbook(template_file)
    if "入力マスター" not in workbook.sheetnames:
        st.error("❌ template.xlsx に『入力マスター』というシートが存在しません。")
        st.stop()

    sheet = workbook["入力マスター"]
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row[1:]:
            cell.value = None

    for idx, row in result_df.iterrows():
        sheet.cell(row=idx+2, column=2, value=row["企業名"])
        sheet.cell(row=idx+2, column=3, value=row["業種"])
        sheet.cell(row=idx+2, column=4, value=row["住所"])
        sheet.cell(row=idx+2, column=5, value=row["電話番号"])

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    st.download_button(
        label="📥 整形済みリストをダウンロード",
        data=output,
        file_name=f"{filename_no_ext}リスト.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
