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
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver4.8 原文電話保持＋入力マスター優先）")

# =========================
# ユーティリティ（正規化系）
# =========================
def nfkc(s: str) -> str:
    return unicodedata.normalize("NFKC", s)

def normalize_text(x) -> str:
    """軽量正規化：NFKC、空白・各種ダッシュ統一、前後空白除去"""
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

# ---- 電話整形系（原文保持と数値キー） ----
# さまざまなハイフンを許容して“原文のまま”抽出するための集合
HYPHENS = "-‒–—―−－ー-‐﹣"
PHONE_TOKEN_RE = re.compile(rf"(\d{{2,4}}[{HYPHENS}]?\d{{2,4}}[{HYPHENS}]?\d{{3,4}})")

def pick_phone_token_raw(line: str) -> str:
    """行から“原文の電話表記”だけを抜き出す（ハイフン位置・種類は一切変更しない）"""
    s = str(line or "")
    m = PHONE_TOKEN_RE.search(s)
    return m.group(1).strip() if m else ""

# normalize_phone は見栄え調整用（原文保持をOFFにしたときだけ使用）
Z2H_HYPHEN = str.maketrans({
    '－':'-','ー':'-','‐':'-','-':'-','‒':'-','–':'-','—':'-','―':'-',
    '-':'-',   # U+2011 non-breaking hyphen
    '−':'-',   # U+2212 minus sign
    '﹣':'-',  # U+FE63 small hyphen-minus
})
def normalize_phone(raw: str) -> str:
    """日本の電話番号をできる限り正確に成形（表示用）。比較は別途 digits で行う。"""
    if not raw:
        return ""
    s = nfkc(str(raw)).translate(Z2H_HYPHEN)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"(内線\d+|内線|内|分機\d+|\(内線.*?\)|\(代\)|\(代表\))", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^\+81", "0", s)
    digits = re.sub(r"\D", "", s)
    if len(digits) < 9:
        return ""
    def fmt(a,b,c): return f"{digits[:a]}-{digits[a:a+b]}-{digits[a+b:a+b+c]}"
    if digits.startswith(("070","080","090","050")) and len(digits) == 11:
        return fmt(3,4,4)
    if digits.startswith("0120") and len(digits) == 10:
        return fmt(4,3,3)
    if len(digits) == 10 and digits.startswith(("03","06")):
        return f"{digits[:2]}-{digits[2:6]}-{digits[6:]}"  # 2-4-4
    if len(digits) == 10:
        return fmt(3,3,4)
    if len(digits) == 11:
        return fmt(3,4,4)
    return re.sub(r"[^\d-]", "", s)

def phone_digits_only(s: str) -> str:
    """比較用の数値キー（NG照合・重複判定はこちら）"""
    return re.sub(r"\D", "", nfkc(str(s or "")))

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
# 入力UI
# =========================
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
selected_nglist = st.selectbox("🛡️ 使用するNGリストを選択してください", nglist_options)

# 電話番号の原文保持（全プロファイルに適用）
keep_original_phone = st.checkbox("電話番号の表記を原文のまま保持する（推奨：Google縦型）", value=True)

st.markdown("### 🧭 抽出方法を選択してください")
profile = st.selectbox(
    "抽出プロファイル",
    [
        "Google検索リスト（縦読み・電話上下型）",
        "シゴトアルワ検索リスト（縦積みラベル）",
        "日本倉庫協会リスト（4列・複数行ブロック）",
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
# 抽出ロジック（3方式＋入力マスター優先）
# =========================
def extract_google_vertical(lines):
    """
    Google縦型：
    - 電話行を見つけたら、直前の“住所らしい行”を取り、その上の“会社名らしい行”に紐付け
    - 電話表示は原文保持オプションに従う
    """
    results = []
    rows = [normalize_text(l) for l in lines if normalize_text(l)]
    address_keywords = ["都","道","府","県","市","区","町","村"]
    company_keywords = ["株式会社","有限会社","合同会社","合名会社","合資会社","(株)","（株）"]
    for i, line in enumerate(rows):
        raw_token = pick_phone_token_raw(line)
        if not raw_token:
            continue
        phone_display = raw_token if keep_original_phone else normalize_phone(raw_token)
        address = ""
        industry = ""
        company = ""
        # 住所探索
        for j in range(i - 1, -1, -1):
            if any(k in rows[j] for k in address_keywords):
                address = rows[j]
                # 会社名探索
                for k in range(j - 1, -1, -1):
                    if any(c in rows[k] for c in company_keywords):
                        company = rows[k]
                        if k + 1 < j:
                            industry = extract_industry(rows[k + 1])
                        break
                break
        results.append([company, industry, address, phone_display])
    return pd.DataFrame(results, columns=["企業名", "業種", "住所", "電話番号"])

def extract_shigoto_arua(df_like: pd.DataFrame) -> pd.DataFrame:
    """左：ラベル/企業名、右：値。電話表示は原文保持オプションに従う。"""
    df = df_like.copy()
    if df.columns.size > 2:
        df = df.iloc[:, :2]
    df.columns = ["col0", "col1"]
    df["col0"] = df["col0"].map(normalize_text)
    df["col1"] = df["col1"].map(lambda x: x if keep_original_phone else normalize_text(x))

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
            phone_val = current["電話番号"]
            # シゴトアルワはセルが電話単体なので原文/整形を切替
            phone_display = str(phone_val).strip() if keep_original_phone else normalize_phone(phone_val)
            out.append([current["企業名"], current["業種"], current["住所"], phone_display])
        current.update({"企業名":"","住所":"","電話番号":"","業種":""})

    for _, row in df.iterrows():
        left = norm_label(row["col0"])
        right = row["col1"]

        # 企業名開始
        if left and (right == "" or right is None) and left not in non_company_labels:
            if current["企業名"]:
                flush_current()
            current["企業名"] = left
            continue

        # ラベル行
        if left in label_to_field and right is not None:
            key = label_to_field[left]
            if key == "住所":
                current["住所"] = clean_address(right)
            elif key == "電話番号":
                current["電話番号"] = right  # 後で原文/整形を切替
            elif key == "業種":
                current["業種"] = extract_industry(right)
            continue

    if current["企業名"]:
        flush_current()

    return pd.DataFrame(out, columns=["企業名", "業種", "住所", "電話番号"])

def extract_warehouse_association(df_like: pd.DataFrame) -> pd.DataFrame:
    """日本倉庫協会：4列ブロック。電話表示は原文保持オプションに従う。"""
    df = df_like.copy()
    if df.shape[1] < 2:
        return pd.DataFrame(columns=["企業名","業種","住所","電話番号"])
    while df.shape[1] < 4:
        df[f"__pad{df.shape[1]}"] = ""
    df = df.iloc[:, :4]
    df.columns = ["c0","c1","c2","c3"]
    for c in df.columns:
        df[c] = df[c].map(normalize_text)

    FACILITY_KEYWORDS = ["営業所","センター","支店","事業所","出張所","倉庫","デポ","物流センター","配送センター"]
    LEGAL_KEYWORDS = ["株式会社","（株）","(株)","有限会社","合同会社","合名会社","合資会社","Inc","INC","Co.,","CO.,","Ltd","LTD","Corp","CORP"]

    def looks_like_company(name: str) -> bool:
        if not name: return False
        if any(k in name for k in FACILITY_KEYWORDS): return False
        if any(k in name for k in LEGAL_KEYWORDS): return True
        return False

    out = []
    current = {"企業名":"", "住所":"", "電話番号":"", "業種_set":set()}

    def flush_current():
        if current["企業名"]:
            industry = "・".join([x for x in current["業種_set"] if x]) or ""
            raw = current["電話番号"]
            phone_display = pick_phone_token_raw(raw) if keep_original_phone else normalize_phone(raw)
            out.append([current["企業名"], industry, current["住所"], phone_display])
        current.update({"企業名":"", "住所":"", "電話番号":"", "業種_set":set()})

    tel_re = re.compile(r"(TEL|ＴＥＬ)\s*([0-9０-９\-ｰー－]+)", re.IGNORECASE)
    zip_re = re.compile(r"^〒\s*\d{3}-?\d{4}")

    for _, row in df.iterrows():
        c0, c1, c2, c3 = row["c0"], row["c1"], row["c2"], row["c3"]

        if c0 and looks_like_company(c0):
            if current["企業名"] and c0 != current["企業名"]:
                flush_current()
            current["企業名"] = c0

        if c1:
            if zip_re.search(c1):
                current["住所"] = c1 if not current["住所"] else f"{current['住所']} {c1}"
            else:
                if any(tok in c1 for tok in ["都","道","府","県","市","区","町","村"]):
                    current["住所"] = c1 if not current["住所"] else f"{current['住所']} {c1}"

        if c2:
            m = tel_re.search(c2)
            if m and not current["電話番号"]:
                current["電話番号"] = m.group(2)

        if c3:
            current["業種_set"].add(extract_industry(c3))

    if current["企業名"]:
        flush_current()

    return pd.DataFrame(out, columns=["企業名","業種","住所","電話番号"])

# =========================
# 共通ユーティリティ
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
    xl = pd.ExcelFile(uploaded_file)

    # === 入力マスター優先（B:企業名/C:業種/D:住所/E:電話） ===
    if "入力マスター" in xl.sheet_names:
        df_raw = pd.read_excel(xl, sheet_name="入力マスター", header=None).fillna("")
        # 電話は原文保持ONなら“そのまま”、OFFなら normalize_phone
        raw_phone_series = df_raw.iloc[:, 4].astype(str)
        disp_phone_series = raw_phone_series.map(lambda v: str(v).strip() if keep_original_phone else normalize_phone(v))
        result_df = pd.DataFrame({
            "企業名": df_raw.iloc[:, 1].astype(str).map(normalize_text),
            "業種": df_raw.iloc[:, 2].astype(str).map(normalize_text),
            "住所": df_raw.iloc[:, 3].astype(str).map(clean_address),
            "電話番号": disp_phone_series
        })
    else:
        # --- 抽出（固定プロファイル） ---
        if profile == "Google検索リスト（縦読み・電話上下型）":
            df = pd.read_excel(uploaded_file, header=None).fillna("")
            lines = df.iloc[:, 0].tolist()
            result_df = extract_google_vertical(lines)

        elif profile == "シゴトアルワ検索リスト（縦積みラベル）":
            df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None).fillna("")
            result_df = extract_shigoto_arua(df0)

        else:  # 日本倉庫協会
            df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None).fillna("")
            result_df = extract_warehouse_association(df0)

    # --- 正規化＆比較キー ---
    result_df = result_df.fillna("")
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

    # --- NGリスト／重複削除／サマリー ---
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
        # NG電話は表記ゆれがあるため digits をキー化
        ng_df["__ng_phone_digits"]  = ng_df.iloc[:, 1].map(phone_digits_only)

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

    # --- Excel出力（物流ハイライトも反映） ---
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
        sheet.cell(row=r, column=5, value=row["電話番号"])  # 表示は原文保持の結果
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
