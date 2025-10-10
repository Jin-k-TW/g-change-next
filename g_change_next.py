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
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver5.0 原文電話保持＋安全配列補正スイッチ＋市外局番監査）")

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

HYPHENS = "-‒–—―−－ー‐﹣\u2011"
Z2H_HYPHEN = str.maketrans({c: "-" for c in HYPHENS})

def phone_digits_only(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))

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

# 🔁 新スイッチ：安全な場合だけ配列（ハイフン位置）補正（原文は保持）
fix_layout_only = st.checkbox(
    "📞 安全な場合のみ配列を自動補正（原文は保持）",
    value=False,
    help="0120/0800/0570/0990/携帯/050/020は規定配列に、固定電話は住所から推定した市外局番で始まる10桁のときだけ〈局番-市内-加入者〉に整えます。数字は一切変更しません。"
)

# =========================
# 抽出ロジック（3方式）※電話は原文保持
# =========================
# 電話らしき「原文部分」をそのまま抜く（変形しない）
HYPHENS_CLASS = re.escape(HYPHENS)
PHONE_TOKEN_RE = re.compile(rf"(\d{{2,4}}(?:[{HYPHENS_CLASS}\s]?\d{{2,4}}){{1,2}})")

def pick_phone_token_raw(line: str) -> str:
    s = str(line or "")
    m = PHONE_TOKEN_RE.search(s)
    return m.group(1).strip() if m else ""

# 1) Google検索リスト
def extract_google_vertical(lines):
    results = []
    rows = [str(l) for l in lines if str(l).strip() != ""]
    for i, line in enumerate(rows):
        ph_raw = pick_phone_token_raw(line)
        if ph_raw:
            phone = ph_raw  # 原文保持
            address = rows[i - 1] if i - 1 >= 0 else ""
            address = clean_address(address)
            industry = extract_industry(rows[i - 2]) if i - 2 >= 0 else ""
            company = rows[i - 3] if i - 3 >= 0 else ""
            results.append([company, industry, address, phone])
    return pd.DataFrame(results, columns=["企業名", "業種", "住所", "電話番号"])

# 2) シゴトアルワ
def extract_shigoto_arua(df_like: pd.DataFrame) -> pd.DataFrame:
    df = df_like.copy()
    if df.columns.size > 2:
        df = df.iloc[:, :2]
    df.columns = ["col0", "col1"]
    df["col0"] = df["col0"].map(lambda x: str(x) if pd.notnull(x) else "")
    df["col1"] = df["col1"].map(lambda x: str(x) if pd.notnull(x) else "")

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
                current["電話番号"]  # 原文保持
            ])
        current["企業名"] = ""
        current["住所"] = ""
        current["電話番号"] = ""
        current["業種"] = ""

    for _, row in df.iterrows():
        left = norm_label(row["col0"])
        right = row["col1"]

        if left and (right == "" or right is None) and left not in non_company_labels:
            if current["企業名"]:
                flush_current()
            current["企業名"] = left
            continue

        if left in label_to_field and right:
            key = label_to_field[left]
            if key == "住所":
                current["住所"] = clean_address(right)
            elif key == "電話番号":
                current["電話番号"] = right  # 原文保持
            elif key == "業種":
                current["業種"] = extract_industry(right)
            continue

    if current["企業名"]:
        flush_current()

    return pd.DataFrame(out, columns=["企業名", "業種", "住所", "電話番号"])

# 3) 日本倉庫協会
def extract_warehouse_association(df_like: pd.DataFrame) -> pd.DataFrame:
    df = df_like.copy()
    if df.shape[1] < 2:
        return pd.DataFrame(columns=["企業名","業種","住所","電話番号"])

    while df.shape[1] < 4:
        df[f"__pad{df.shape[1]}"] = ""
    df = df.iloc[:, :4]
    df.columns = ["c0","c1","c2","c3"]
    for c in df.columns:
        df[c] = df[c].map(lambda x: str(x) if pd.notnull(x) else "")

    FACILITY_KEYWORDS = ["営業所","センター","支店","事業所","出張所","倉庫","デポ","物流センター","配送センター"]
    LEGAL_KEYWORDS = ["株式会社","（株）","(株)","有限会社","合同会社","合名会社","合資会社","Inc","INC","Co.,","CO.,","Ltd","LTD","Corp","CORP"]

    def looks_like_company(name: str) -> bool:
        if not name:
            return False
        if any(k in name for k in FACILITY_KEYWORDS):
            return False
        if any(k in name for k in LEGAL_KEYWORDS):
            return True
        return False

    out = []
    current = {"企業名":"", "住所":"", "電話番号":"", "業種_set":set()}

    def flush_current():
        if current["企業名"]:
            industry = "・".join([x for x in current["業種_set"] if x]) or ""
            out.append([
                current["企業名"],
                industry,
                current["住所"],
                current["電話番号"]  # 原文保持
            ])
        current["企業名"] = ""
        current["住所"] = ""
        current["電話番号"] = ""
        current["業種_set"] = set()

    tel_re = re.compile(r"(?:TEL|Tel|tel)\s*([0-9０-９\-ｰー－\s]+)")
    zip_re = re.compile(r"^〒\s*\d{3}-?\d{4}")

    for _, row in df.iterrows():
        c0, c1, c2, c3 = row["c0"], row["c1"], row["c2"], row["c3"]

        if c0 and looks_like_company(c0):
            if current["企業名"] and c0 != current["企業名"]:
                flush_current()
            current["企業名"] = c0

        if c1:
            if zip_re.search(c1):
                if not current["住所"]:
                    current["住所"] = c1
                elif c1 not in current["住所"]:
                    current["住所"] = f"{current['住所']} {c1}".strip()
            else:
                if any(tok in c1 for tok in ["都","道","府","県","市","区","町","村"]):
                    if current["住所"]:
                        if c1 not in current["住所"]:
                            current["住所"] = f"{current['住所']} {c1}".strip()
                    else:
                        current["住所"] = c1

        if c2:
            m = tel_re.search(c2)
            if m and not current["電話番号"]:
                current["電話番号"] = m.group(1).strip()  # 原文保持

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

    # === 入力マスター優先（電話は原文保持）===
    if "入力マスター" in xl.sheet_names:
        df_raw = pd.read_excel(xl, sheet_name="入力マスター", header=None).fillna("")
        result_df = pd.DataFrame({
            "企業名": df_raw.iloc[:, 1].astype(str),
            "業種": df_raw.iloc[:, 2].astype(str),
            "住所": df_raw.iloc[:, 3].astype(str),
            "電話番号": df_raw.iloc[:, 4].astype(str)  # 原文保持
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

    # --- 正規化＆比較キー（会社名正規化・電話digits） ---
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

    # --- NGリスト／重複削除／サマリー（現状維持：電話はdigits照合のみ） ---
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
        ng_df["__ng_phone_digits"]  = ng_df.iloc[:, 1].astype(str).map(phone_digits_only)

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
                    "source_phone": row["電話番号"],  # 原文をログ
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
                "source_phone": row["電話番号"],  # 原文をログ
                "match_key": row["__phone_digits"],
                "ng_hit": ""
            })
        result_df = result_df[~dup_mask]
    removed_by_dedup = before - len(result_df)

    # --- 住所→市区町村→市外局番 監査（原文は上書きしない） ---
    AREACODE_CSV = "jp_areacodes.csv"      # 列: prefecture,municipality,area_code
    TOWN2CITY_CSV = "jp_town2city.csv"     # 列: prefecture,municipality,town_keyword

    PREFS = ["北海道","青森県","岩手県","宮城県","秋田県","山形県","福島県","茨城県","栃木県","群馬県","埼玉県","千葉県","東京都","神奈川県","新潟県","富山県","石川県","福井県","山梨県","長野県","岐阜県","静岡県","愛知県","三重県","滋賀県","京都府","大阪府","兵庫県","奈良県","和歌山県","鳥取県","島根県","岡山県","広島県","山口県","徳島県","香川県","愛媛県","高知県","福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県","沖縄県"]

    def detect_pref(address: str) -> str:
        s = normalize_text(address)
        for p in PREFS:
            if p in s:
                return p
        return ""

    def load_town2city(path=TOWN2CITY_CSV):
        if not os.path.exists(path):
            return pd.DataFrame(columns=["prefecture","municipality","town_keyword"])
        return pd.read_csv(path, dtype=str).fillna("")

    T2C = load_town2city()

    def detect_city(pref: str, address: str) -> str:
        if not pref or T2C.empty:
            return ""
        s = normalize_text(address)
        cand = T2C[T2C["prefecture"] == pref]
        hits = []
        for _, r in cand.iterrows():
            tk = str(r["town_keyword"])
            if tk and tk in s:
                hits.append((len(tk), r["municipality"]))
        if hits:
            hits.sort(reverse=True)
            return hits[0][1]
        for muni in cand["municipality"].unique():
            if str(muni) in s:
                return str(muni)
        return ""

    def load_areacodes(path=AREACODE_CSV):
        if not os.path.exists(path):
            return pd.DataFrame(columns=["prefecture","municipality","area_code"])
        return pd.read_csv(path, dtype=str).fillna("")

    AC = load_areacodes()

    def guess_areacode(pref: str, muni: str) -> str:
        if AC.empty or not pref:
            return ""
        if muni:
            m = AC[(AC["prefecture"]==pref) & (AC["municipality"]==muni)]
            if not m.empty:
                return str(m.iloc[0]["area_code"])
        m2 = AC[(AC["prefecture"]==pref) & (AC["municipality"].isin(["","-","なし","_","NA"]))]
        if not m2.empty:
            return str(m2.iloc[0]["area_code"])
        return ""

    def starts_with_areacode(digits: str, ac: str) -> bool:
        return bool(digits and ac and digits.startswith(ac))

    audit_rows = []
    for idx, row in result_df.iterrows():
        addr = row.get("住所","")
        raw  = row.get("電話番号","")
        digits = phone_digits_only(raw)
        pref = detect_pref(addr)
        muni = detect_city(pref, addr)
        ac   = guess_areacode(pref, muni)
        ok   = starts_with_areacode(digits, ac) if ac else None

        result_df.loc[idx, "__addr_pref"]  = pref
        result_df.loc[idx, "__addr_city"]  = muni
        result_df.loc[idx, "__suggest_ac"] = ac
        result_df.loc[idx, "__ac_match"]   = ok

        audit_rows.append({
            "企業名": row.get("企業名",""),
            "住所": addr,
            "電話番号(原文保持)": raw,
            "番号digits": digits,
            "推定都道府県": pref,
            "推定市区町村": muni,
            "推定市外局番": ac,
            "市外局番一致": "" if ok is None else ("一致" if ok else "不一致")
        })

    mismatch_cnt = sum(1 for r in audit_rows if r["市外局番一致"]=="不一致")

    # --- 表示用電話番号（原文を尊重。スイッチONなら安全に配列のみ整形） ---
    def format_special_or_mobile(digits: str) -> str | None:
        # フリーダイヤル等
        if digits.startswith("0120") and len(digits) == 10:
            return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"      # 4-3-3
        if digits.startswith("0800") and len(digits) == 11:
            return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"      # 4-3-4
        if (digits.startswith("0570") or digits.startswith("0990")) and len(digits) == 10:
            return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"      # 4-3-3
        # 携帯/050/020
        if len(digits) == 11 and digits.startswith(("070","080","090","050","020")):
            return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"      # 3-4-4
        return None

    def format_fixed_with_areacode(digits: str, ac: str) -> str | None:
        """固定10桁で、住所から推定した市外局番acで始まる場合のみ、配列を〈局番-市内-加入者4〉にする"""
        if not ac or len(digits) != 10 or not digits.startswith(ac):
            return None
        mid_len = (10 - len(ac) - 4)
        if mid_len <= 0:
            return None
        return f"{ac}-{digits[len(ac):len(ac)+mid_len]}-{digits[-4:]}"

    display_numbers = []
    fixed_count = 0
    for _, row in result_df.iterrows():
        raw = row.get("電話番号","")
        digits = phone_digits_only(raw)
        disp = raw  # 基本は原文

        if fix_layout_only and digits:
            s = format_special_or_mobile(digits)
            if s:
                disp = s
                fixed_count += 1
            else:
                ac = (row.get("__suggest_ac") or "").strip()
                s2 = format_fixed_with_areacode(digits, ac)
                if s2:
                    disp = s2
                    fixed_count += 1

        display_numbers.append(disp)

    result_df["__display_phone"] = display_numbers

    # --- 空行除去・並べ替え（空電話は最後／表示は原文or補正表示） ---
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
        st.dataframe(
            result_df[["企業名","業種","住所","__display_phone"]].rename(columns={"__display_phone":"電話番号"}),
            use_container_width=True
        )

    # --- サマリー＋削除ログDL＋市外局番監査 ---
    with st.expander("📊 実行サマリー（詳細）"):
        st.markdown(f"""
- フィルター除外（製造業 完全一致＋一部部分一致）: **{removed_by_industry}** 件  
- NG（企業名・部分一致）削除: **{company_removed}** 件  
- NG（電話・数字一致）削除: **{phone_removed}** 件  
- 重複（電話・数字一致）削除: **{removed_by_dedup}** 件  
- 市外局番の不一致（住所推定と番号digitsの先頭が異なる）: **{mismatch_cnt}** 件  
- 配列の自動補正（スイッチON時のみ）: **{fixed_count}** 件  
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
        audit_df = pd.DataFrame(audit_rows)
        st.dataframe(audit_df.head(50), use_container_width=True)
        audit_csv = audit_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "🔎 市外局番 監査CSVをダウンロード",
            data=audit_csv,
            file_name="area_code_audit.csv",
            mime="text/csv"
        )

    # --- Excel出力（電話は表示用の列を出力／物流ハイライトも反映） ---
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
        sheet.cell(row=r, column=5, value=row["__display_phone"])  # 原文 or 安全補正
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
    st.info("template.xlsx / jp_areacodes.csv / jp_town2city.csv と（必要なら）NGリストxlsxを同じフォルダに置いてから、Excelファイルをアップロードしてください。")
