import streamlit as st
import pandas as pd
import re
import io
import os
import unicodedata
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =========================
# ページ設定／スタイル
# =========================
st.set_page_config(page_title="G-Change Next", layout="wide")
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver4.6 + 縦積み詳細プロファイル）")

# =========================
# ユーティリティ（正規化系）
# =========================
def nfkc(s: str) -> str:
    return unicodedata.normalize("NFKC", s)

def normalize_text(x) -> str:
    """共通の軽量正規化：NFKC、空白・各種ダッシュ統一、前後空白除去"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).replace("\u3000", " ").replace("\xa0", " ")
    s = re.sub(r'[−–—―ー]', '-', s)
    s = nfkc(s).strip()
    return s

def hiragana_to_katakana(s: str) -> str:
    """ひらがな→カタカナ（翻字ではない。英字⇄カナ/漢字の相互変換は行わない）"""
    res = []
    for ch in s:
        code = ord(ch)
        if 0x3041 <= code <= 0x3096:  # ひらがな範囲
            res.append(chr(code + 0x60))  # カタカナへ
        else:
            res.append(ch)
    return "".join(res)

COMPANY_SUFFIXES = [
    "株式会社", "(株)", "（株）", "有限会社", "(有)", "（有）",
    "inc.", "inc", "co.,ltd.", "co.,ltd", "co.ltd.", "co.ltd", "ltd.", "ltd",
    "corp.", "corp", "co.", "co",
    "合同会社", "合名会社", "合資会社"
]

def canonical_company_name(name: str) -> str:
    """
    企業名の比較用キー（強い正規化）
    - NFKC 等（normalize_text）
    - ひら→カナ統一（※翻字はしない）
    - 英字の大小無視（casefold）
    - 会社種別語の除去
    - 記号・空白を比較用に削る
    """
    s = normalize_text(name)
    s = hiragana_to_katakana(s)
    s = s.casefold()
    for suf in sorted(COMPANY_SUFFIXES, key=len, reverse=True):
        s = s.replace(suf.casefold(), "")
    s = re.sub(r"[\s\-–—―‐ー・/,.·･\(\)（）\[\]{}【】＆&＋+_|]", "", s)
    return s

Z2H_HYPHEN = str.maketrans({
    '－':'-','ー':'-','‐':'-','-':'-','‒':'-','–':'-','—':'-','―':'-'
})

def normalize_phone(raw: str) -> str:
    """表示用の軽い整形（比較は phone_digits_only() を使用）"""
    if not raw:
        return ""
    s = nfkc(raw).translate(Z2H_HYPHEN)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"(\(内線.*?\)|\(代\)|\(代表\))", "", s)
    s = re.sub(r"^\+81", "0", s)
    s = re.sub(r"[^\d-]", "", s)
    digits = re.sub(r"\D", "", s)
    if len(digits) < 9:
        return ""
    if len(digits) == 10:
        return f"{digits[0:3]}-{digits[3:6]}-{digits[6:]}"
    if len(digits) == 11:
        return f"{digits[0:3]}-{digits[3:7]}-{digits[7:]}"
    return s

def phone_digits_only(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def clean_address(address):
    address = normalize_text(address)
    split_pattern = r"[·･・]"
    if re.search(split_pattern, address):
        return re.split(split_pattern, address)[-1].strip()
    return address

def extract_industry(line):
    parts = re.split(r"[·・]", normalize_text(line))
    return parts[-1].strip() if len(parts) > 1 else (normalize_text(line))

# =========================
# 既存フィルター定義（現状維持）
# =========================
remove_exact = [
    "オフィス機器レンタル業", "足場レンタル会社", "電気工", "廃棄物リサイクル業",
    "プロパン販売業者", "看板専門店", "給水設備工場", "警備業", "建設会社",
    "工務店", "写真店", "人材派遣業", "整備店", "倉庫", "肉店", "米販売店",
    "スーパーマーケット", "ロジスティクスサービス", "建材店",
    "自動車整備工場", "自動車販売店", "車体整備店", "協会/組織", "建設請負業者", "電器店", "家電量販店", "建築会社", "ハウス クリーニング業", "焼肉店",
    "建築設計事務所","左官","作業服店","空調設備工事業者","金属スクラップ業者","害獣駆除サービス","モーター修理店","アーチェリーショップ","アスベスト検査業","事務用品店",
    "測量士","配管業者","労働組合","ガス会社","ガソリンスタンド","ガラス/ミラー店","ワイナリー","屋根ふき業者","高等学校","金物店","史跡","商工会議所","清掃業","清掃業者","配管工"
]
remove_partial = ["販売店", "販売業者"]

highlight_partial = [
    "運輸", "ロジスティクスサービス", "倉庫", "輸送サービス",
    "運送会社企業のオフィス", "運送会社"
]

# =========================
# 入力UI（既存 + 抽出プロファイルを前段に追加）
# =========================
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
selected_nglist = st.selectbox("🛡️ 使用するNGリストを選択してください", nglist_options)

st.markdown("### 🧭 抽出プロファイルを選択してください")
profile = st.selectbox(
    "抽出プロファイル",
    ["自動判定（おすすめ）", "縦積み詳細（ラベル付き）", "従来：入力マスター/1列縦"]
)

st.markdown("### 🏭 業種カテゴリを選択してください")
industry_option = st.radio(
    "どの業種カテゴリーに該当しますか？",
    ("製造業", "物流業", "その他")
)

uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロード", type=["xlsx"])

# =========================
# 抽出ロジック
# =========================
def extract_company_groups_legacy(lines):
    """（従来）電話らしい行を基準に 企業名/業種/住所/電話 を拾う簡易ヒューリスティック"""
    results = []
    rows = [normalize_text(l) for l in lines if normalize_text(l)]
    for i, line in enumerate(rows):
        ph = normalize_phone(line)
        if ph:
            phone = ph
            address = rows[i - 1] if i - 1 >= 0 else ""
            address = clean_address(address)
            industry = extract_industry(rows[i - 2]) if i - 2 >= 0 else ""
            company = rows[i - 3] if i - 3 >= 0 else ""
            results.append([company, industry, address, phone])
    return pd.DataFrame(results, columns=["企業名", "業種", "住所", "電話番号"])

def extract_vertical_labeled(df_like: pd.DataFrame) -> pd.DataFrame:
    """
    新規：縦積み詳細（ラベル付き）形式を抽出。
    想定：2列（左：ラベルor企業名、右：値）。企業名行は右がNaN/空で、左に社名。
    以降「住所」「電話」「業種」などのラベル行が続く。
    """
    # ヘッダー行をデータとして扱うため header=None 推奨
    df = df_like.copy()
    if df.columns.size > 2:
        # 2列超のときは最初の2列だけを見る（安全サイド）
        df = df.iloc[:, :2]
    df.columns = ["col0", "col1"]
    df["col0"] = df["col0"].map(normalize_text)
    df["col1"] = df["col1"].map(normalize_text)

    label_candidates = {"住所": "住所", "電話": "電話番号", "TEL": "電話番号", "Tel": "電話番号", "tel": "電話番号", "業種": "業種"}
    current = {"企業名": "", "住所": "", "電話番号": "", "業種": ""}
    out = []

    def flush_current():
        if any(current.values()) and current["企業名"]:
            out.append([current["企業名"], current["業種"], current["住所"], normalize_phone(current["電話番号"])])
        # リセット
        current["企業名"] = ""
        current["住所"] = ""
        current["電話番号"] = ""
        current["業種"] = ""

    for _, row in df.iterrows():
        left = row["col0"]
        right = row["col1"]

        # 企業名の開始条件：右が空で、左が非空、かつラベル語でない
        if left and (right == "" or right is None) and left not in label_candidates.keys():
            # 既に積んでいるものがあれば確定
            if current["企業名"]:
                flush_current()
            current["企業名"] = left
            continue

        # ラベル行
        if left in label_candidates:
            key = label_candidates[left]
            if key == "住所":
                current["住所"] = clean_address(right)
            elif key == "電話番号":
                current["電話番号"] = right
            elif key == "業種":
                current["業種"] = extract_industry(right)
            continue

        # それ以外のラベルは無視（資本金やFAX等）

    # 最終行 flush
    if current["企業名"]:
        flush_current()

    return pd.DataFrame(out, columns=["企業名", "業種", "住所", "電話番号"])

def auto_detect_and_extract(xl: pd.ExcelFile) -> pd.DataFrame:
    """
    自動判定：最初のシートを軽く見て、縦積み詳細っぽければその抽出、
    それ以外は従来ロジック（入力マスター or 1列縦）へ。
    """
    sheet_names = xl.sheet_names
    # まず「入力マスター」優先（従来互換）
    if "入力マスター" in sheet_names:
        df_raw = pd.read_excel(xl, sheet_name="入力マスター", header=None).fillna("")
        return pd.DataFrame({
            "企業名": df_raw.iloc[:, 1].astype(str).map(normalize_text),
            "業種": df_raw.iloc[:, 2].astype(str).map(normalize_text),
            "住所": df_raw.iloc[:, 3].astype(str).map(clean_address),
            "電話番号": df_raw.iloc[:, 4].astype(str).map(normalize_phone)
        })

    # 先頭シートを header=None で読んで、縦積み判定
    df0 = pd.read_excel(xl, sheet_name=sheet_names[0], header=None).fillna("")
    # 縦積み判定：2列以上 かつ 左列に「住所/電話/業種」ラベルが頻出
    left_values = df0.iloc[:, 0].astype(str).tolist()
    label_hits = sum(v in ["住所", "電話", "TEL", "Tel", "tel", "業種"] for v in left_values)
    if df0.shape[1] >= 2 and label_hits >= 2:
        return extract_vertical_labeled(df0.iloc[:, :2])

    # それ以外は従来：1列縦→4行セット抽出
    lines = df0.iloc[:, 0].tolist()
    return extract_company_groups_legacy(lines)

def clean_dataframe(df):
    return df.fillna("").applymap(lambda x: normalize_text(x) if pd.notnull(x) else "")

def remove_empty_rows(df):
    return df[~((df["企業名"] == "") & (df["業種"] == "") & (df["住所"] == "") & (df["電話番号"] == ""))]

# =========================
# メイン処理
# =========================
if uploaded_file:
    filename_no_ext = os.path.splitext(uploaded_file.name)[0]
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names

    # ---- 抽出（プロファイルに応じて） ----
    if profile == "自動判定（おすすめ）":
        result_df = auto_detect_and_extract(xl)

    elif profile == "縦積み詳細（ラベル付き）":
        # 明示指定：先頭シートを header=None で読み、縦積み抽出
        df0 = pd.read_excel(xl, sheet_name=sheet_names[0], header=None).fillna("")
        result_df = extract_vertical_labeled(df0)

    else:  # "従来：入力マスター/1列縦"
        if "入力マスター" in sheet_names:
            df_raw = pd.read_excel(uploaded_file, sheet_name="入力マスター", header=None).fillna("")
            result_df = pd.DataFrame({
                "企業名": df_raw.iloc[:, 1].astype(str).map(normalize_text),
                "業種": df_raw.iloc[:, 2].astype(str).map(normalize_text),
                "住所": df_raw.iloc[:, 3].astype(str).map(clean_address),
                "電話番号": df_raw.iloc[:, 4].astype(str).map(normalize_phone)
            })
        else:
            df = pd.read_excel(uploaded_file, header=None).fillna("")
            lines = df.iloc[:, 0].tolist()
            result_df = extract_company_groups_legacy(lines)

    # ---- 正規化（現状維持） ----
    result_df = clean_dataframe(result_df)
    # 比較用キー
    result_df["__company_canon"] = result_df["企業名"].map(canonical_company_name)
    result_df["__phone_digits"]  = result_df["電話番号"].map(phone_digits_only)

    # ---- 業種フィルター（現状維持） ----
    removed_by_industry = 0
    styled_df = None
    if industry_option == "製造業":
        before = len(result_df)
        result_df = result_df[~result_df["業種"].isin(remove_exact)]
        if remove_partial:
            pat = "|".join(map(re.escape, remove_partial))
            result_df = result_df[~result_df["業種"].str.contains(pat, na=False)]
        removed_by_industry = before - len(result_df)
        st.warning(f"🏭 製造業フィルター適用：{removed_by_industry}件を除外しました")

    elif industry_option == "物流業":
        def highlight_logistics(val):
            v = val or ""
            return "background-color: red" if any(word in v for word in highlight_partial) else ""
        styled_df = result_df.style.applymap(highlight_logistics, subset=["業種"])
        st.info("🚚 業種が一致したセルを赤くハイライトしています（出力にも反映）")

    # ---- NGリスト適用（会社名=部分一致／電話=数字一致）＋削除ログ（現状維持） ----
    removal_logs = []
    company_removed = 0
    phone_removed = 0

    if selected_nglist != "なし":
        ng_path = f"{selected_nglist}.xlsx"
        if not os.path.exists(ng_path):
            st.error(f"❌ 選択されたNGリストファイルが見つかりません：{ng_path}")
            st.stop()
        ng_df = pd.read_excel(ng_path).fillna("")
        if ng_df.shape[1] < 2:
            st.error("❌ NGリストは2列以上必要です（企業名、電話番号）")
            st.stop()

        ng_df["__ng_company_canon"] = ng_df.iloc[:, 0].map(canonical_company_name)
        ng_df["__ng_phone_digits"]  = ng_df.iloc[:, 1].astype(str).map(normalize_phone).map(phone_digits_only)

        ng_company_keys = ng_df["__ng_company_canon"].tolist()
        ng_phone_set    = set([p for p in ng_df["__ng_phone_digits"].tolist() if p])

        # 企業名（canonical部分一致）
        before = len(result_df)

        def hit_ng_company(canon_name: str):
            for ng in ng_company_keys:
                if ng and canon_name and (ng in canon_name or canon_name in ng):
                    return ng
            return None

        hit_indices = []
        for idx, row in result_df.iterrows():
            ng_key = hit_ng_company(row["__company_canon"])
            if ng_key:
                removal_logs.append({
                    "reason": "ng-company",
                    "source_company": row["企業名"],
                    "source_phone": row["電話番号"],
                    "match_key": row["__company_canon"],
                    "ng_hit": ng_key
                })
                hit_indices.append(idx)
        if hit_indices:
            result_df = result_df.drop(index=hit_indices)
        company_removed = before - len(result_df)

        # 電話（数字一致）
        before = len(result_df)
        hits = result_df["__phone_digits"].isin(ng_phone_set)
        if hits.any():
            for idx, row in result_df[hits].iterrows():
                removal_logs.append({
                    "reason": "ng-phone",
                    "source_company": row["企業名"],
                    "source_phone": row["電話番号"],
                    "match_key": row["__phone_digits"],
                    "ng_hit": row["__phone_digits"]
                })
            result_df = result_df[~hits]
        phone_removed = before - len(result_df)

    # ---- 重複削除：電話（数字一致）のみ（現状維持） ----
    before = len(result_df)
    dup_mask = result_df["__phone_digits"].ne("").astype(bool) & result_df["__phone_digits"].duplicated(keep="first")
    if dup_mask.any():
        for idx, row in result_df[dup_mask].iterrows():
            removal_logs.append({
                "reason": "phone-duplicate",
                "source_company": row["企業名"],
                "source_phone": row["電話番号"],
                "match_key": row["__phone_digits"],
                "ng_hit": ""
            })
        result_df = result_df[~dup_mask]
    removed_by_dedup = before - len(result_df)

    # ---- 空行除去・並べ替え（現状維持） ----
    result_df = remove_empty_rows(result_df)
    result_df["_phdigits"] = result_df["__phone_digits"]
    result_df["_is_empty_phone"] = (result_df["_phdigits"] == "")
    result_df = result_df.sort_values(by=["_is_empty_phone", "_phdigits", "企業名"]).drop(columns=["_phdigits","_is_empty_phone"])
    result_df = result_df.reset_index(drop=True)

    # ---- 画面表示（現状維持） ----
    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")
    if industry_option == "物流業" and styled_df is not None:
        st.dataframe(styled_df, use_container_width=True)
    else:
        st.dataframe(result_df[["企業名","業種","住所","電話番号"]], use_container_width=True)

    # ---- サマリー＋削除ログDL（現状維持） ----
    with st.expander("📊 実行サマリー（詳細）"):
        st.markdown(f"""
- フィルター除外（製造業 完全一致＋一部部分一致）: **{removed_by_industry}** 件  
- NG（企業名・部分一致）削除: **{company_removed}** 件  
- NG（電話・数字一致）削除: **{phone_removed}** 件  
- 重複（電話・数字一致）削除: **{removed_by_dedup}** 件  
""")
        if removal_logs:
            log_df = pd.DataFrame(removal_logs)
            st.dataframe(log_df.head(100), use_container_width=True)
            csv_bytes = log_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "🧾 削除ログをCSVでダウンロード",
                data=csv_bytes,
                file_name="removal_logs.csv",
                mime="text/csv"
            )

    # ---- Excel出力（現状維持：物流ハイライトも反映） ----
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
        for cell in row[1:5]:
            cell.value = None
            cell.fill = PatternFill(fill_type=None)

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    def is_logi(val: str) -> bool:
        v = val or ""
        return any(word in v for word in highlight_partial)

    for idx, row in result_df.iterrows():
        r = idx + 2
        sheet.cell(row=r, column=2, value=row["企業名"])
        sheet.cell(row=r, column=3, value=row["業種"])
        sheet.cell(row=r, column=4, value=row["住所"])
        sheet.cell(row=r, column=5, value=row["電話番号"])
        if industry_option == "物流業" and is_logi(row["業種"]):
            sheet.cell(row=r, column=3).fill = red_fill

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    st.download_button(
        label="📥 整形済みリストをダウンロード",
        data=output,
        file_name=f"{filename_no_ext}リスト.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("template.xlsx と（必要なら）NGリストxlsxを同じフォルダに置いてから、Excelファイルをアップロードしてください。")
