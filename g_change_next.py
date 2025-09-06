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
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver4.7.1 固定プロファイル＋物流協会修正）")

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
    res = []
    for ch in s:
        code = ord(ch)
        if 0x3041 <= code <= 0x3096:
            res.append(chr(code + 0x60))
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
    """表示用の軽い整形（比較は phone_digits_only を使用）"""
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
# 入力UI（固定プロファイル）
# =========================
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
selected_nglist = st.selectbox("🛡️ 使用するNGリストを選択してください", nglist_options)

st.markdown("### 🧭 抽出方法を選択してください")
profile = st.selectbox(
    "抽出プロファイル",
    [
        "Google検索リスト（縦読み・電話上下型）",
        "シゴトアルワ検索リスト（縦積みラベル）",
        "物流協会リスト（4列・複数行ブロック）",
    ],
    index=0
)

st.markdown("### 🏭 業種カテゴリを選択してください")
industry_option = st.radio(
    "どの業種カテゴリーに該当しますか？",
    ("製造業", "物流業", "その他")
)

uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロード", type=["xlsx"])

# =========================
# 抽出ロジック（3方式）
# =========================
# 1) Google検索リスト：1列縦。電話行を軸に、上3行を 企業名/業種/住所 とみなす既存方式
def extract_google_vertical(lines):
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

# 2) シゴトアルワ：2列縦。左がラベル/企業名、右が値
def extract_shigoto_arua(df_like: pd.DataFrame) -> pd.DataFrame:
    df = df_like.copy()
    if df.columns.size > 2:
        df = df.iloc[:, :2]
    df.columns = ["col0", "col1"]
    df["col0"] = df["col0"].map(normalize_text)
    df["col1"] = df["col1"].map(normalize_text)

    def norm_label(s: str) -> str:
        s = (s or "")
        s = re.sub(r"[：:]\s*$", "", s)
        return s

    label_to_field = {
        "住所": "住所",
        "所在地": "住所",
        "本社所在地": "住所",
        "電話": "電話番号",
        "電話番号": "電話番号",
        "TEL": "電話番号",
        "Tel": "電話番号",
        "tel": "電話番号",
        "業種": "業種",
        "事業内容": "業種",
        "産業分類": "業種",
        "製造業種": "業種",
    }

    non_company_labels = set([
        "住所","所在地","本社所在地",
        "電話","電話番号","TEL","Tel","tel",
        "FAX","ＦＡＸ",
        "資本金","資本金（千円）","資本金(千円)",
        "従業員数","設立年月",
        "業種","事業内容","産業分類","製造業種"
    ])

    current = {"企業名": "", "住所": "", "電話番号": "", "業種": ""}
    out = []

    def flush_current():
        if current["企業名"]:
            out.append([
                current["企業名"],
                current["業種"],
                current["住所"],
                normalize_phone(current["電話番号"])
            ])
        current["企業名"] = ""
        current["住所"] = ""
        current["電話番号"] = ""
        current["業種"] = ""

    for _, row in df.iterrows():
        left = norm_label(row["col0"])
        right = row["col1"]

        # 企業名開始（右が空・左がラベル語でない）
        if left and (right == "" or right is None) and left not in non_company_labels:
            if current["企業名"]:
                flush_current()
            current["企業名"] = left
            continue

        # ラベル行（値あり）
        if left in label_to_field and right:
            key = label_to_field[left]
            if key == "住所":
                current["住所"] = clean_address(right)
            elif key == "電話番号":
                current["電話番号"] = right
            elif key == "業種":
                current["業種"] = extract_industry(right)
            continue

    if current["企業名"]:
        flush_current()

    return pd.DataFrame(out, columns=["企業名", "業種", "住所", "電話番号"])

# 3) 物流協会：4列×複数行ブロック（★会社開始を厳格化）
def extract_butsuryu_association(df_like: pd.DataFrame) -> pd.DataFrame:
    """
    想定：
      c0: 会社名 or 施設名（営業所/センター等） or 空
      c1: 郵便(〒) or 住所本体
      c2: 'TEL xxx' / 'FAX xxx'
      c3: 倉庫種別など（複数行）
    施設名は会社開始にしない（直前会社の情報として扱う）
    """
    df = df_like.copy()
    if df.shape[1] < 2:
        return pd.DataFrame(columns=["企業名","業種","住所","電話番号"])

    # 4列に合わせる
    while df.shape[1] < 4:
        df[f"__pad{df.shape[1]}"] = ""
    df = df.iloc[:, :4]
    df.columns = ["c0","c1","c2","c3"]
    for c in df.columns:
        df[c] = df[c].map(normalize_text)

    # 会社開始判定ルール
    FACILITY_KEYWORDS = ["営業所","センター","支店","事業所","出張所","倉庫","デポ","物流センター","配送センター"]
    LEGAL_KEYWORDS = ["株式会社","（株）","(株)","有限会社","合同会社","合名会社","合資会社","Inc","INC","Co.,","CO.,","Ltd","LTD","Corp","CORP"]

    def looks_like_company(name: str) -> bool:
        """法人らしい表記を含むなら True（施設ワードが含まれていたら False）"""
        if not name:
            return False
        if any(k in name for k in FACILITY_KEYWORDS):
            return False
        if any(k in name for k in LEGAL_KEYWORDS):
            return True
        return False  # 安全側：法人表記が無ければ会社開始にしない

    out = []
    current = {"企業名":"", "住所":"", "電話番号":"", "業種_set":set()}

    def flush_current():
        if current["企業名"]:
            industry = "・".join([x for x in current["業種_set"] if x]) or ""
            out.append([
                current["企業名"],
                industry,
                current["住所"],
                normalize_phone(current["電話番号"])
            ])
        current["企業名"] = ""
        current["住所"] = ""
        current["電話番号"] = ""
        current["業種_set"] = set()

    tel_re = re.compile(r"TEL\s*([0-9０-９\-ｰー－]+)", re.IGNORECASE)
    zip_re = re.compile(r"^〒\s*\d{3}-?\d{4}")

    for _, row in df.iterrows():
        c0, c1, c2, c3 = row["c0"], row["c1"], row["c2"], row["c3"]

        # --- 会社開始（法人らしい時のみ）---
        if c0 and looks_like_company(c0):
            if current["企業名"] and c0 != current["企業名"]:
                flush_current()
            current["企業名"] = c0

        # --- 住所（〒＋住所本体の連結）---
        if c1:
            if zip_re.search(c1):
                if not current["住所"]:
                    current["住所"] = c1
                elif c1 not in current["住所"]:
                    current["住所"] = f"{current['住所']} {c1}".strip()
            else:
                # 住所らしいキーワード（都/道/府/県/市/区/町/村）を含む場合のみ結合
                if any(tok in c1 for tok in ["都","道","府","県","市","区","町","村"]):
                    if current["住所"]:
                        if c1 not in current["住所"]:
                            current["住所"] = f"{current['住所']} {c1}".strip()
                    else:
                        current["住所"] = c1

        # --- 電話（最初のひとつ）---
        if c2:
            m = tel_re.search(c2)
            if m and not current["電話番号"]:
                current["電話番号"] = m.group(1)

        # --- 業種断片（あれば集約）---
        if c3:
            current["業種_set"].add(extract_industry(c3))

    # 最終 flush
    if current["企業名"]:
        flush_current()

    return pd.DataFrame(out, columns=["企業名","業種","住所","電話番号"])

# =========================
# 後段の共通関数
# =========================
def clean_dataframe(df):
    return df.fillna("").applymap(lambda x: normalize_text(x) if pd.notnull(x) else "")

def remove_empty_rows(df):
    return df[~((df["企業名"] == "") & (df["業種"] == "") & (df["住所"] == "") & (df["電話番号"] == ""))]

# =========================
# メイン処理
# =========================
if uploaded_file:
    filename_no_ext = os.path.splitext(uploaded_file.name)[0]

    # --- 抽出（固定プロファイル） ---
    if profile == "Google検索リスト（縦読み・電話上下型）":
        df = pd.read_excel(uploaded_file, header=None).fillna("")
        lines = df.iloc[:, 0].tolist()
        result_df = extract_google_vertical(lines)

    elif profile == "シゴトアルワ検索リスト（縦積みラベル）":
        xl = pd.ExcelFile(uploaded_file)
        df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None).fillna("")
        result_df = extract_shigoto_arua(df0)

    else:  # 物流協会
        xl = pd.ExcelFile(uploaded_file)
        df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None).fillna("")
        result_df = extract_butsuryu_association(df0)

    # --- 正規化＆比較キー ---
    result_df = clean_dataframe(result_df)
    result_df["__company_canon"] = result_df["企業名"].map(canonical_company_name)
    result_df["__phone_digits"]  = result_df["電話番号"].map(phone_digits_only)

    # --- 業種フィルター（現状維持） ---
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

    # --- NGリスト／重複削除／サマリー（現状維持） ---
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

    # --- 空行除去・並べ替え（空電話は最後） ---
    result_df = remove_empty_rows(result_df)
    result_df["_phdigits"] = result_df["__phone_digits"]
    result_df["_is_empty_phone"] = (result_df["_phdigits"] == "")
    result_df = result_df.sort_values(by=["_is_empty_phone", "_phdigits", "企業名"]).drop(columns=["_phdigits","_is_empty_phone"])
    result_df = result_df.reset_index(drop=True)

    # --- 画面表示 ---
    st.success(f"✅ 整形完了：{len(result_df)}件の企業データを取得しました。")
    if industry_option == "物流業" and styled_df is not None:
        st.dataframe(styled_df, use_container_width=True)
    else:
        st.dataframe(result_df[["企業名","業種","住所","電話番号"]], use_container_width=True)

    # --- サマリー＋削除ログDL ---
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

    # --- Excel出力（現状維持：物流ハイライトも反映） ---
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
