import streamlit as st
import pandas as pd
import re
import io
import os
import unicodedata
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ページ設定
st.set_page_config(page_title="G-Change Next", layout="wide")

# デザイン設定
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver4.4 強化版）")

# --- NGリスト選択 ---
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
selected_nglist = st.selectbox("🛡️ 使用するNGリストを選択してください", nglist_options)

# --- 業種フィルター選択 ---
st.markdown("### 🏭 業種カテゴリを選択してください")
industry_option = st.radio(
    "どの業種カテゴリーに該当しますか？",
    ("製造業", "物流業", "その他")
)

# フィルター定義
remove_exact = [
    "オフィス機器レンタル業", "足場レンタル会社", "電気工", "廃棄物リサイクル業",
    "プロパン販売業者", "看板専門店", "給水設備工場", "警備業", "建設会社",
    "工務店", "写真店", "人材派遣業", "整備店", "倉庫", "肉店", "米販売店",
    "スーパーマーケット", "ロジスティクスサービス", "建材店",
    "自動車整備工場", "自動車販売店", "車体整備店", "協会/組織", "建設請負業者", "電器店"
]
remove_partial = ["販売店", "販売業者"]

highlight_partial = [
    "運輸", "ロジスティクスサービス", "倉庫", "輸送サービス",
    "運送会社企業のオフィス", "運送会社"
]

# --- ファイルアップロード ---
uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロード", type=["xlsx"])

# --- ユーティリティ関数群 ---
def normalize(text):
    if text is None or pd.isna(text):
        return ""
    text = str(text).strip().replace("\u3000", " ").replace("\xa0", " ")
    text = re.sub(r'[−–—―]', '-', text)
    text = unicodedata.normalize("NFKC", text)
    return text

def clean_address(address):
    address = normalize(address)
    split_pattern = r"[·･・]"
    if re.search(split_pattern, address):
        return re.split(split_pattern, address)[-1].strip()
    return address

def extract_phone(line):
    match = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)
    return match.group() if match else ""

def extract_industry(line):
    parts = re.split(r"[·・]", line)
    return parts[-1].strip() if len(parts) > 1 else line.strip()

def extract_company_groups(lines):
    results = []
    lines = [normalize(l) for l in lines if l and normalize(l)]
    for i, line in enumerate(lines):
        if extract_phone(line):
            phone = extract_phone(line)
            address = lines[i - 1] if i - 1 >= 0 else ""
            industry = extract_industry(lines[i - 2]) if i - 2 >= 0 else ""
            company = lines[i - 3] if i - 3 >= 0 else ""
            results.append([company, industry, address, phone])
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
            "住所": df_raw.iloc[:, 3].astype(str).apply(clean_address),
            "電話番号": df_raw.iloc[:, 4].astype(str).apply(normalize)
        })
    else:
        df = pd.read_excel(uploaded_file, header=None)
        lines = df[0].dropna().tolist()
        result_df = extract_company_groups(lines)

    result_df = clean_dataframe(result_df)

    # --- 業種フィルター処理 ---
    if industry_option == "製造業":
        before = len(result_df)
        result_df = result_df[~result_df["業種"].isin(remove_exact)]
        result_df = result_df[~result_df["業種"].str.contains("|".join(remove_partial), na=False)]
        st.warning(f"🏭 製造業フィルター適用：{before - len(result_df)}件を除外しました")

    elif industry_option == "物流業":
        def highlight_logistics(val):
            if any(word in val for word in highlight_partial):
                return "background-color: red"
            return ""
        styled_df = result_df.style.applymap(highlight_logistics, subset=["業種"])
        st.info("🚚 業種が一致したセルを赤くハイライトしています（出力にも反映）")
    else:
        styled_df = result_df

    # --- NGリスト処理 ---
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
        ng_companies = ng_df.iloc[:, 0].dropna().astype(str).apply(normalize).tolist()
        ng_phones = ng_df.iloc[:, 1].dropna().astype(str).apply(normalize).tolist()

        before_company = len(result_df)
        result_df = result_df[~result_df["企業名"].apply(lambda x: any(ng in normalize(x) for ng in ng_companies))]
        company_removed = before_company - len(result_df)

        before_phone = len(result_df)
        result_df = result_df[~result_df["電話番号"].apply(normalize).isin(ng_phones)]
        phone_removed = before_phone - len(result_df)

    result_df = remove_phone_duplicates(result_df)
    result_df = remove_empty_rows(result_df)
    result_df = result_df.sort_values(by="電話番号", na_position='last').reset_index(drop=True)

    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")

    if industry_option == "物流業":
        st.dataframe(styled_df, use_container_width=True)
    else:
        st.dataframe(result_df, use_container_width=True)

    if selected_nglist != "なし":
        st.info(f"🛡️ 【NGリスト削除件数】\n\n企業名による削除：{company_removed}件\n電話番号による削除：{phone_removed}件")

    # --- Excel出力処理 ---
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

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for idx, row in result_df.iterrows():
        sheet.cell(row=idx+2, column=2, value=row["企業名"])
        sheet.cell(row=idx+2, column=3, value=row["業種"])
        sheet.cell(row=idx+2, column=4, value=row["住所"])
        sheet.cell(row=idx+2, column=5, value=row["電話番号"])
        if industry_option == "物流業":
            if any(word in row["業種"] for word in highlight_partial):
                sheet.cell(row=idx+2, column=3).fill = red_fill

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    st.download_button(
        label="📥 整形済みリストをダウンロード",
        data=output,
        file_name=f"{filename_no_ext}リスト.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
