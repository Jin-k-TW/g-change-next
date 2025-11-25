import streamlit as st
import pandas as pd
import re
import unicodedata
import io
import os
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ===============================
# 簡易ログイン（パスワード認証）
# ===============================
def check_password():
    """st.secrets['password'] と一致するかを確認する簡易ログイン"""
    def password_entered():
        """テキストボックスに入力されたパスワードを検証"""
        if "password" not in st.secrets:
            st.session_state["password_correct"] = False
            st.session_state["password_error"] = "サーバー側にパスワードが設定されていません。管理者に確認してください。"
            return

        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            st.session_state.pop("password", None)
            st.session_state.pop("password_error", None)
        else:
            st.session_state["password_correct"] = False
            st.session_state["password_error"] = "パスワードが違います。もう一度入力してください。"

    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔐 G-Change Next ログイン")
        st.text_input(
            "パスワードを入力してください",
            type="password",
            on_change=password_entered,
            key="password",
        )
        if "password_error" in st.session_state and st.session_state["password_error"]:
            st.error(st.session_state["password_error"])
        return False

    return True

# ===============================
# Streamlit設定
# ===============================
st.set_page_config(page_title="G-Change Next", layout="wide")

if not check_password():
    st.stop()

st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver6.3 複数ファイル対応＋確定ボタン省略版）")

# ===============================
# テキスト正規化
# ===============================
def nfkc(s: str) -> str:
    return unicodedata.normalize("NFKC", s)

def normalize_text(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).replace("\u3000", " ").replace("\xa0", " ")
    s = re.sub(r'[−–—―ー]', '-', s)
    return nfkc(s).strip()

def clean_address(address: str) -> str:
    address = normalize_text(address)
    return address.strip()

def extract_industry(line: str) -> str:
    return normalize_text(line)

# ===============================
# 企業名正規化（NG照合用）
# ===============================
COMPANY_SUFFIXES = ["株式会社", "(株)", "（株）", "有限会社", "(有)", "（有）", "合同会社"]
def canonical_company_name(name: str) -> str:
    s = normalize_text(name)
    for suf in sorted(COMPANY_SUFFIXES, key=len, reverse=True):
        s = s.replace(suf, "")
    s = re.sub(r"[\s\-・/,.·･\(\)（）【】＆&＋+_|]", "", s)
    return s

# ===============================
# 電話番号処理（原文保持）
# ===============================
HYPHENS = "-‒–—―−－ー‐﹣\u2011"
HYPHENS_CLASS = re.escape(HYPHENS)

CANDIDATE_RE = re.compile(rf"[+]?\d(?:[\d{HYPHENS_CLASS}\s]{{6,}})\d")

def pick_phone_token_raw(line: str) -> str:
    """1行から電話番号らしい文字列を抽出。digits 長が 9〜11 以外は不採用。"""
    if not line:
        return ""
    s = unicodedata.normalize("NFKC", str(line))
    raw_cands = CANDIDATE_RE.findall(s)
    cands = []
    for token in raw_cands:
        tok = token.strip()
        if ":" in tok:
            continue
        digits = re.sub(r"\D", "", tok)
        if not (9 <= len(digits) <= 11):
            continue
        if not (digits.startswith("0") or digits.startswith("81")):
            continue
        score = (len(digits), tok.count("-"))
        cands.append((score, tok))
    if not cands:
        return ""
    cands.sort(key=lambda x: x[0], reverse=True)
    return cands[0][1]

def phone_digits_only(s: str) -> str:
    return re.sub(r"\D", "", str(s or ""))

# ===============================
# 住所判定＋業種分割ロジック（強化版）
# ===============================

JAPAN_PREFS = [
    "北海道",
    "青森県","岩手県","宮城県","秋田県","山形県","福島県",
    "茨城県","栃木県","群馬県","埼玉県","千葉県","東京都","神奈川県",
    "新潟県","富山県","石川県","福井県","山梨県","長野県",
    "岐阜県","静岡県","愛知県","三重県",
    "滋賀県","京都府","大阪府","兵庫県","奈良県","和歌山県",
    "鳥取県","島根県","岡山県","広島県","山口県",
    "徳島県","香川県","愛媛県","高知県",
    "福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県","沖縄県",
]

def split_industry_and_address(line: str):
    """
    1セルに「業種 + 住所」が入っているケースを想定。
    左側 = 業種, 右側 = 住所 として分割する。

    住所の開始位置は次の順で探す：
      1) 都道府県名
      2) 「市 / 区 / 郡 / 町 / 村」など
      3) 郵便番号 (123-4567 / 〒123-4567)
      4) 最初の数字
    見つからない場合は (\"\", \"\") を返して旧ロジックにフォールバック。
    """
    if not line:
        return ("", "")

    s = normalize_text(line)

    addr_pos = None

    # 1) 都道府県名
    for pref in JAPAN_PREFS:
        idx = s.find(pref)
        if idx != -1 and (addr_pos is None or idx < addr_pos):
            addr_pos = idx

    # 2) 市/区/郡/町/村 など（都道府県が見つからなかったとき）
    if addr_pos is None:
        m_city = re.search(r".{0,5}(市|区|郡|町|村)", s)
        if m_city:
            addr_pos = max(0, m_city.start())

    # 3) 郵便番号 (123-4567 / 〒123-4567)
    if addr_pos is None:
        m_zip = re.search(r"〒?\d{3}-\d{4}", s)
        if m_zip:
            addr_pos = m_zip.start()

    # 4) 最初の数字（番地・丁目など）
    if addr_pos is None:
        m_num = re.search(r"\d", s)
        if m_num:
            addr_pos = m_num.start()

    # それでも見つからない → 住所とは判断しない（旧ロジックに任せる）
    if addr_pos is None:
        return ("", "")

    industry = s[:addr_pos].strip(" ・:：　")
    address  = s[addr_pos:].strip()

    # 住所候補があまりに住所っぽくない場合はフォールバック
    if not address:
        return ("", "")

    # 市区町村 or 数字が一切無い場合は住所とみなさない
    if not re.search(r"[市区郡町村]", address) and not re.search(r"\d", address):
        return ("", "")

    return (industry, address)

# ===============================
# 抽出プロファイル（3方式）
# ===============================

# 1) Google検索リスト（縦読み・電話上下）
def extract_google_vertical(lines):
    """
    Googleマップ縦型リスト用。
    - パターンA（新）：電話行の1つ上に「業種+住所」のセル、その上に企業名
    - パターンB（旧）：電話行の1つ上に住所、2つ上に業種、3つ上に企業名
    の両方を自動で判定して処理する。
    """
    results = []
    rows = [str(l) for l in lines if str(l).strip() != ""]

    for i, line in enumerate(rows):
        ph_raw = pick_phone_token_raw(line)
        if not ph_raw:
            continue

        phone = ph_raw
        company = ""
        industry = ""
        address = ""

        # 直上の行を業種+住所セルとして試す（パターンA）
        mid = rows[i - 1] if i - 1 >= 0 else ""
        ind_a, addr_a = split_industry_and_address(mid)

        if addr_a:  # 住所が判定できた → 新パターン
            address = clean_address(addr_a)
            industry = ind_a  # 空ならあとで補完
            company = rows[i - 2] if i - 2 >= 0 else ""

            # 業種が空なら、さらに1つ上の行を業種として補完してみる
            if not industry and i - 2 >= 0:
                industry = extract_industry(rows[i - 2])

        else:
            # 旧パターンにフォールバック
            address = clean_address(mid)
            industry = extract_industry(rows[i - 2]) if i - 2 >= 0 else ""
            company = rows[i - 3] if i - 3 >= 0 else ""

        results.append([company, industry, address, phone])

    return pd.DataFrame(results, columns=["企業名", "業種", "住所", "電話番号"])

# 2) シゴトアルワ（縦積み）
def extract_shigoto_arua(df_like: pd.DataFrame) -> pd.DataFrame:
    df = df_like.copy()
    if df.columns.size > 2:
        df = df.iloc[:, :2]
    df.columns = ["col0", "col1"]
    df = df.fillna("")
    current = {"企業名": "", "住所": "", "電話番号": "", "業種": ""}
    out = []

    def flush():
        if current["企業名"]:
            out.append([current["企業名"], current["業種"], current["住所"], current["電話番号"]])
        current.update({"企業名": "", "住所": "", "電話番号": "", "業種": ""})

    for _, row in df.iterrows():
        k, v = str(row["col0"]), str(row["col1"])
        if k in ["住所", "所在地", "本社所在地"]:
            current["住所"] = clean_address(v)
        elif k in ["電話", "電話番号", "TEL", "Tel", "tel"]:
            current["電話番号"] = v
        elif k in ["業種", "事業内容", "産業分類", "製造業種"]:
            current["業種"] = extract_industry(v)
        elif k and not v:
            if current["企業名"]:
                flush()
            current["企業名"] = k
    if current["企業名"]:
        flush()
    return pd.DataFrame(out, columns=["企業名", "業種", "住所", "電話番号"])

# 3) 日本倉庫協会（4列）
def extract_warehouse_association(df_like: pd.DataFrame) -> pd.DataFrame:
    df = df_like.fillna("")
    if df.shape[1] < 2:
        return pd.DataFrame(columns=["企業名", "業種", "住所", "電話番号"])
    while df.shape[1] < 4:
        df[f"__pad{df.shape[1]}"] = ""
    df = df.iloc[:, :4]
    df.columns = ["c0", "c1", "c2", "c3"]

    tel_re = re.compile(r"(?:TEL|Tel|tel)\s*([0-9０-９\-ー－\s]+)")
    out, current = [], {"企業名": "", "住所": "", "電話番号": "", "業種_set": set()}

    def flush():
        if current["企業名"]:
            out.append([current["企業名"], "・".join(current["業種_set"]), current["住所"], current["電話番号"]])
        current.update({"企業名": "", "住所": "", "電話番号": "", "業種_set": set()})

    for _, r in df.iterrows():
        if r["c0"]:
            if current["企業名"] and r["c0"] != current["企業名"]:
                flush()
            current["企業名"] = r["c0"]
        if r["c1"]:
            current["住所"] = clean_address(r["c1"])
        if r["c2"]:
            m = tel_re.search(r["c2"])
            if m:
                current["電話番号"] = m.group(1).strip()
        if r["c3"]:
            current["業種_set"].add(extract_industry(r["c3"]))
    if current["企業名"]:
        flush()
    return pd.DataFrame(out, columns=["企業名", "業種", "住所", "電話番号"])

# ===============================
# 業種のフィルター/ハイライト
# ===============================
remove_exact = [
    "オフィス機器レンタル業", "足場レンタル会社", "電気工", "廃棄物リサイクル業",
    "プロパン販売業者", "看板専門店", "給水設備工場", "警備業", "建設会社",
    "工務店", "写真店", "人材派遣業", "整備店", "倉庫", "肉店", "米販売店",
    "スーパーマーケット", "ロジスティクスサービス", "建材店",
    "自動車整備工場", "自動車販売店", "車体整備店", "協会/組織", "建設請負業者", "電器店", "家電量販店",
    "建築会社", "ハウス クリーニング業", "焼肉店", "建築設計事務所", "左官",
    "作業服店", "空調設備工事業者", "金属スクラップ業者", "害獣駆除サービス",
    "モーター修理店", "アーチェリーショップ", "アスベスト検査業", "事務用品店",
    "測量士", "配管業者", "労働組合", "ガス会社", "ガソリンスタンド",
    "ガラス/ミラー店", "ワイナリー", "屋根ふき業者", "高等学校", "金物店",
    "史跡", "商工会議所", "清掃業", "清掃業者", "配管工", "お手頃"
]
remove_partial = ["販売店", "販売業者"]

highlight_partial = [
    "運輸", "ロジスティクスサービス", "倉庫", "輸送サービス",
    "運送会社企業のオフィス", "運送会社"
]

# ===============================
# 業種ノイズ除去（レビュー/評価など）
# ===============================
def clean_industry_noise(s: str) -> str:
    """
    業種カラムに紛れ込むレビュー系ノイズを削除。
    ＋ 最後に「·」「レビュ-なし」「空白だけ」は必ず消す。
    """
    if not s:
        return ""
    t = str(s)
    t = re.sub(r"\s+", " ", t).strip()

    t = re.sub(r"^\s*\d+(?:\.\d+)?\s*[\(（]\s*\d+\s*[\)）]\s*(?:件)?\s*[・･]?\s*", "", t)

    def norm_token(x: str) -> str:
        return re.sub(r"\s+", "", x)

    noise_basic = {"レビュー", "レビューなし", "レビュー無し", "クチコミ", "口コミ"}
    noise_nashi = {"なし"}

    if t.startswith("レビュー"):
        parts = [p.strip() for p in re.split(r"[・･]", t) if p.strip()]
        if not parts:
            return ""
        if all(norm_token(p) in noise_basic | noise_nashi for p in parts):
            return ""
        cleaned_parts = []
        for p in parts:
            pn = norm_token(p)
            if pn in noise_basic or pn in noise_nashi:
                continue
            cleaned_parts.append(p)
        t = "・".join(cleaned_parts)
    else:
        t = re.sub(r"(?:^|[・･])\s*(Google\s*の?\s*クチコミ|口コミ|クチコミ)\s*(?=[・･]|$)", "", t)
        t = re.sub(r"[・･]?\s*\d+\s*件の?(レビュー|口コミ|クチコミ)\s*(?=[・･]|$)", "", t)

    parts = [p.strip() for p in re.split(r"[・･]", t) if p.strip()]
    t = "・".join(parts) if parts else ""
    t = re.sub(r"[・･]{2,}", "・", t).strip(" ・･")

    if t:
        for trash in ["·", "レビュ-なし"]:
            t = t.replace(trash, "")
        t = re.sub(r"\s+", " ", t).strip()

    return t if t else ""

# ===============================
# 共通整形（電話は触らない）
# ===============================
def clean_dataframe_except_phone(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ["企業名", "業種", "住所"]:
        df[c] = df[c].map(normalize_text)
    df["業種"] = df["業種"].map(clean_industry_noise)
    return df.fillna("")

# ===============================
# UI
# ===============================
st.markdown("### 🛡️ 使用するNGリストを選択")
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "NGリスト" in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
selected_nglist = st.selectbox(
    "NGリスト",
    nglist_options,
    index=0,
    help="同じフォルダにある『NGリスト〜.xlsx』を検出します。1列目=企業名、2列目=電話番号（任意）。"
)

st.markdown("### 🧭 抽出方法を選択")
profile = st.selectbox(
    "抽出プロファイル",
    ["Google検索リスト（縦読み・電話上下型）", "シゴトアルワ検索リスト（縦積み）", "日本倉庫協会リスト（4列型）"]
)

st.markdown("### 🏭 業種カテゴリを選択")
industry_option = st.radio("どの業種カテゴリーに該当しますか？", ("製造業", "物流業", "その他"))

st.markdown("### 🧩 テンプレートの取得方法（OS互換強化）")
template_source = st.radio(
    "template.xlsx の取得元",
    ("プロジェクト内の template.xlsx を使う（従来）", "ここで template.xlsx をアップロードして使う"),
    index=0
)
template_upload = None
if template_source == "ここで template.xlsx をアップロードして使う":
    template_upload = st.file_uploader("template.xlsx をアップロード", type=["xlsx"], key="template_up")

uploaded_files = st.file_uploader(
    "📤 整形対象のExcelファイルをアップロード（複数選択可）",
    type=["xlsx"],
    accept_multiple_files=True
)

# ===============================
# NGリスト読み込み
# ===============================
ng_names = []
ng_phones = set()
if uploaded_files and selected_nglist != "なし":
    ng_path = f"{selected_nglist}.xlsx"
    if not os.path.exists(ng_path):
        st.error(f"❌ 選択されたNGリストが見つかりません：{ng_path}")
        st.stop()
    ng_df = pd.read_excel(ng_path, engine="openpyxl").fillna("")
    if ng_df.shape[1] < 1:
        st.error("❌ NGリストは少なくとも1列（企業名）が必要です。2列目に電話番号があれば照合に利用します。")
        st.stop()
    ng_df["__ng_company_canon"] = ng_df.iloc[:, 0].map(canonical_company_name)
    if ng_df.shape[1] >= 2:
        ng_df["__ng_digits"] = ng_df.iloc[:, 1].astype(str).map(phone_digits_only)
    else:
        ng_df["__ng_digits"] = ""
    ng_names = [n for n in ng_df["__ng_company_canon"].tolist() if n]
    ng_phones = set([d for d in ng_df["__ng_digits"].tolist() if d])

# ===============================
# メイン処理
# ===============================
if uploaded_files:
    for file_index, uploaded_file in enumerate(uploaded_files):
        st.markdown("---")
        st.markdown(f"## 📁 {uploaded_file.name}")

        filename_no_ext = os.path.splitext(uploaded_file.name)[0]
        xl = pd.ExcelFile(uploaded_file, engine="openpyxl")

        # --- 抽出 ---
        if "入力マスター" in xl.sheet_names:
            df_raw = pd.read_excel(
                xl,
                sheet_name="入力マスター",
                header=None,
                engine="openpyxl"
            ).fillna("")
            df = pd.DataFrame({
                "企業名": df_raw.iloc[1:, 1].astype(str),
                "業種": df_raw.iloc[1:, 2].astype(str),
                "住所": df_raw.iloc[1:, 3].astype(str),
                "電話番号": df_raw.iloc[1:, 4].astype(str),
            })
        else:
            if profile == "Google検索リスト（縦読み・電話上下型）":
                df0 = pd.read_excel(uploaded_file, header=None, engine="openpyxl").fillna("")
                lines = df0.iloc[:, 0].tolist()
                df = extract_google_vertical(lines)
            elif profile == "シゴトアルワ検索リスト（縦積み）":
                df0 = pd.read_excel(xl, header=None, engine="openpyxl").fillna("")
                df = extract_shigoto_arua(df0)
            else:
                df0 = pd.read_excel(xl, header=None, engine="openpyxl").fillna("")
                df = extract_warehouse_association(df0)

        df = clean_dataframe_except_phone(df)

        df["__company_canon"] = df["企業名"].map(canonical_company_name)
        df["__digits"] = df["電話番号"].map(phone_digits_only)

        removed_by_industry = 0
        if industry_option == "製造業":
            before = len(df)
            all_ng_words = remove_exact + remove_partial
            if all_ng_words:
                pat = "|".join(map(re.escape, all_ng_words))
                df = df[~df["業種"].str.contains(pat, na=False)]
            removed_by_industry = before - len(df)
            st.warning(f"🏭 製造業フィルター適用：{removed_by_industry}件を除外しました")

        removal_logs = []
        company_removed = 0
        phone_removed = 0
        dup_removed = 0

        if ng_names or ng_phones:
            before = len(df)
            hit_idx = []
            for idx, row in df.iterrows():
                c = row["__company_canon"]
                if not c:
                    continue
                if any((n in c or c in n) for n in ng_names):
                    removal_logs.append({
                        "reason": "ng-company",
                        "company": row["企業名"],
                        "phone_raw": row["電話番号"],
                        "match": c
                    })
                    hit_idx.append(idx)
            if hit_idx:
                df = df.drop(index=hit_idx)
            company_removed = before - len(df)

            before = len(df)
            mask = df["__digits"].isin(ng_phones)
            if mask.any():
                for idx, row in df[mask].iterrows():
                    removal_logs.append({
                        "reason": "ng-phone",
                        "company": row["企業名"],
                        "phone_raw": row["電話番号"],
                        "match": row["__digits"]
                    })
                df = df[~mask]
            phone_removed = before - len(df)

        before = len(df)
        dup_mask = df["__digits"].ne("").astype(bool) & df["__digits"].duplicated(keep="first")
        if dup_mask.any():
            for idx, row in df[dup_mask].iterrows():
                removal_logs.append({
                    "reason": "dup-phone",
                    "company": row["企業名"],
                    "phone_raw": row["電話番号"],
                    "match": row["__digits"]
                })
            df = df[~dup_mask]
        dup_removed = before - len(df)

        df = df[~((df["企業名"] == "") & (df["業種"] == "") & (df["住所"] == "") & (df["電話番号"] == ""))].reset_index(drop=True)

        st.success(f"✅ 整形完了：{len(df)}件の企業データを取得しました。")
        edited = st.data_editor(
            df[["企業名", "業種", "住所", "電話番号"]],
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "企業名": st.column_config.TextColumn(required=True),
                "業種": st.column_config.TextColumn(),
                "住所": st.column_config.TextColumn(),
                "電話番号": st.column_config.TextColumn(
                    help="原文の配列を保持。必要ならここで手動修正してください。編集内容はそのまま出力に反映されます。"
                ),
            },
            key=f"editable_preview_{file_index}",
        )

        df_export = edited.copy()

        with st.expander(f"📊 実行サマリー（詳細） - {uploaded_file.name}", expanded=False):
            st.markdown(
                f"- フィルター除外（製造業 部分一致）: **{removed_by_industry}** 件\n"
                f"- NG（企業名 部分一致）削除: **{company_removed}** 件\n"
                f"- NG（電話 digits一致）削除: **{phone_removed}** 件\n"
                f"- 重複（電話 digits一致）削除: **{dup_removed}** 件\n"
            )
            if removal_logs:
                log_df = pd.DataFrame(removal_logs)
                st.dataframe(log_df.head(300), use_container_width=True)
                csv_bytes = log_df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "🧾 削除ログをCSVでダウンロード",
                    data=csv_bytes,
                    file_name=f"removal_logs_{filename_no_ext}.csv",
                    mime="text/csv",
                    key=f"removal_log_btn_{file_index}",
                )

        # ===============================
        # template.xlsx へ書き込み
        # ===============================
        wb = None
        if template_upload is not None:
            try:
                buf = io.BytesIO(template_upload.read())
                wb = load_workbook(buf)
            except Exception as e:
                st.error(f"❌ アップロードした template.xlsx の読み込みに失敗しました: {e}")
                st.stop()
        else:
            app_dir = Path(__file__).resolve().parent
            template_path = app_dir / "template.xlsx"
            if not template_path.exists():
                st.error(
                    f"❌ template.xlsx が見つかりませんでした（期待パス: {template_path}）。"
                    "『ここで template.xlsx をアップロードして使う』を選ぶか、"
                    "ファイルをプロジェクト直下に配置してください。"
                )
                st.stop()
            try:
                wb = load_workbook(template_path)
            except Exception as e:
                st.error(f"❌ template.xlsx の読み込みに失敗しました: {e}")
                st.stop()

        if "入力マスター" not in wb.sheetnames:
            st.error("❌ template.xlsx に『入力マスター』というシートが存在しません。")
            st.stop()

        sheet = wb["入力マスター"]

        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            for cell in row[1:5]:
                cell.value = None
                cell.fill = PatternFill(fill_type=None)

        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        def is_logi(val: str) -> bool:
            v = (val or "").strip()
            return any(word in v for word in highlight_partial)

        for idx_row, row in df_export.iterrows():
            r = idx_row + 2
            sheet.cell(row=r, column=2, value=row["企業名"])
            sheet.cell(row=r, column=3, value=row["業種"])
            sheet.cell(row=r, column=4, value=row["住所"])
            sheet.cell(row=r, column=5, value=row["電話番号"])
            if industry_option == "物流業" and is_logi(row["業種"]):
                sheet.cell(row=r, column=3).fill = red_fill

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button(
            label=f"📥 整形済みリストをダウンロード（{filename_no_ext} / template.xlsx 反映）",
            data=output,
            file_name=f"{filename_no_ext}リスト.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_btn_{file_index}",
        )

else:
    st.info("Excelファイルをアップロードしてください。NGリストxlsxは同フォルダに置くか、プロジェクト直下に配置してください。")
