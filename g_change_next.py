import streamlit as st
import pandas as pd
import os
import re
import io

# ページ設定
st.set_page_config(page_title="G-Change｜Googleリスト整形＋NGリスト除外", layout="wide")

# タイトル＆スタイル
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G-Change Plus｜Googleリスト整形＋NGリスト除外（入力マスター対応版）")

# ファイルアップロード
uploaded_file = st.file_uploader("📤 整形対象のリストをアップロードしてください", type=["xlsx"])

# NGリストの選択肢を取得
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "リスト" not in f and "template" not in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# 正規化関数
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    text = re.sub(r'[−–—―]', '-', text)
    return text

# 企業情報の抽出（縦型リスト用）
def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""

    for line in lines[1:]:
        line = normalize(line)
        if "·" in line or "⋅" in line:
            parts = re.split(r"[·⋅]", line)
            industry = parts[-1].strip()
        elif re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line):
            phone = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line).group()
        elif not address and any(x in line for x in ["丁目", "町", "番", "区", "−", "-"]):
            address = line

    return pd.Series([company, industry, address, phone])

# 縦型リストかチェック
def is_company_line(line):
    line = normalize(line)
    return not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

# 電話番号重複削除
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

# 空白行除去
def remove_empty_rows(df):
    return df[~((df["企業名"] == "") & (df["業種"] == "") & (df["住所"] == "") & (df["電話番号"] == ""))]

# メイン処理
if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names

    # 入力マスターシートがあるか確認
    if "入力マスター" in sheet_names:
        df_raw = pd.read_excel(uploaded_file, sheet_name="入力マスター", header=None)

        # B〜E列（1〜4列目）を読み込む（ヘッダー不要）
        df = pd.DataFrame({
            "企業名": df_raw.iloc[:, 1].astype(str).apply(normalize),   # B列
            "業種": df_raw.iloc[:, 2].astype(str).apply(normalize),     # C列
            "住所": df_raw.iloc[:, 3].astype(str).apply(normalize),     # D列
            "電話番号": df_raw.iloc[:, 4].astype(str).apply(normalize)  # E列
        })

    else:
        # 縦型リストとして処理
        df_raw = pd.read_excel(uploaded_file, header=None)
        lines = df_raw[0].dropna().tolist()

        groups = []
        current = []
        for line in lines:
            line = normalize(str(line))
            if is_company_line(line):
                if current:
                    groups.append(current)
                current = [line]
            else:
                current.append(line)
        if current:
            groups.append(current)

        df = pd.DataFrame([extract_info(group) for group in groups],
                          columns=["企業名", "業種", "住所", "電話番号"])

    df = df.applymap(lambda x: str(x).strip() if pd.notnull(x) else x)

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")

    # NGリスト除外処理
    if nglist_choice != "なし":
        ng_file_path = nglist_choice + ".xlsx"
        ng_df = pd.read_excel(ng_file_path)

        ng_companies = ng_df["企業名"].dropna().tolist() if "企業名" in ng_df.columns else []
        ng_phones = ng_df["電話番号"].dropna().tolist() if "電話番号" in ng_df.columns else []

        mask_company = df["企業名"].apply(lambda x: any(ng in str(x) for ng in ng_companies))
        mask_phone = df["電話番号"].apply(lambda x: str(x) in [str(p) for p in ng_phones])

        removed_df = df[mask_company | mask_phone]
        df = df[~(mask_company | mask_phone)]

        company_removed = mask_company.sum()
        phone_removed = mask_phone.sum()

        st.success(f"🧹 NGリスト除外完了！（企業名除外：{company_removed}件、電話番号除外：{phone_removed}件）")

        if not removed_df.empty:
            st.error("🚫 除外された企業一覧")
            st.dataframe(removed_df, use_container_width=True)

    df = remove_phone_duplicates(df)
    df = remove_empty_rows(df)
    df = df.sort_values(by="電話番号", na_position='last').reset_index(drop=True)

    # 出力ファイル名
    uploaded_filename = uploaded_file.name.replace(".xlsx", "")
    final_filename = uploaded_filename + "：リスト.xlsx"

    # 保存用バッファ
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="リスト")

    # ダウンロード
    st.download_button(
        label="📥 整形済みリストをダウンロード",
        data=output.getvalue(),
        file_name=final_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
