# 🚗 G-Change Next Ver3.2

import streamlit as st
import pandas as pd
import re
import io
import os

# ページ設定
st.set_page_config(page_title="G-Change Next", layout="wide")

# デザイン設定
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)

# タイトル
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver3.2）")

# --- NGリスト選択ブロック ---

# GitHub直下にある「NGリスト」という名前を含むxlsxファイルを取得
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]

# プルダウンリスト作成
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]

# プルダウン選択
selected_nglist = st.selectbox("🛡️ 使用するNGリストを選択してください", nglist_options)

# --- 整形対象ファイルアップロードブロック ---

uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロード", type=["xlsx"])

# --- テキスト整形ルール ---

review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

def normalize(text):
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r'[−–—―]', '-', text)

def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""

    for line in lines[1:]:
        line = normalize(line)
        if any(kw in line for kw in ignore_keywords):
            continue
        if any(kw in line for kw in review_keywords):
            continue
        if "·" in line or "⋅" in line:
            parts = re.split(r"[·⋅]", line)
            industry = parts[-1].strip()
        elif re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line):
            phone = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line).group()
        elif not address and any(x in line for x in ["丁目", "町", "番", "区", "−", "-"]):
            address = line

    return pd.Series([company, industry, address, phone])

def is_company_line(line):
    line = normalize(str(line))
    return not any(kw in line for kw in ignore_keywords + review_keywords) and not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

# --- 実行メインブロック ---

if uploaded_file:
    # ファイルを読み込む（まずはシート名一覧取得）
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names

    # 「入力マスター」シートがあればテンプレ型と判定
    if "入力マスター" in sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name="入力マスター")
        # 列名で「企業様名称」「業種」「住所」「電話番号」を抜き出し
        if all(col in df.columns for col in ["企業様名称", "業種", "住所", "電話番号"]):
            result_df = df[["企業様名称", "業種", "住所", "電話番号"]].copy()
            result_df.columns = ["企業名", "業種", "住所", "電話番号"]
        else:
            st.error("⚠️ 入力マスターシートに必要な列（企業様名称、業種、住所、電話番号）がありません。")
            st.stop()
    else:
        # 縦型リストパターン
        df = pd.read_excel(uploaded_file, header=None)
        lines = df[0].dropna().tolist()

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

        result_df = pd.DataFrame([extract_info(group) for group in groups],
                                 columns=["企業名", "業種", "住所", "電話番号"])

    # --- NGリスト適用処理 ---
    if selected_nglist != "なし":
        nglist_df = pd.read_excel(f"{selected_nglist}.xlsx")

        ng_companies = nglist_df.iloc[:, 0].dropna().astype(str).tolist()
        ng_phones = nglist_df.iloc[:, 1].dropna().astype(str).tolist()

        # 部分一致（企業名）フィルタ
        result_df = result_df[~result_df["企業名"].apply(lambda x: any(ng_name in str(x) for ng_name in ng_companies))]

        # 完全一致（電話番号）フィルタ
        result_df = result_df[~result_df["電話番号"].astype(str).isin(ng_phones)]

    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")
    st.dataframe(result_df, use_container_width=True)

    # --- Excel保存 ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name="整形済みデータ")
    st.download_button("📥 整形済みExcelファイルをダウンロード", data=output.getvalue(),
                       file_name="整形済み_企業リスト.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")