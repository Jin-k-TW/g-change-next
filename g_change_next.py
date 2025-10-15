import streamlit as st
import pandas as pd
import re
import unicodedata
import io
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ===============================
# Streamlit設定
# ===============================
st.set_page_config(page_title="G-Change Next", layout="wide")
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver5.2 原文電話保持＋誤検出防止）")

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

# 電話番号候補抽出（誤検出防止）
CANDIDATE_RE = re.compile(
    rf"[+]?\d(?:[\d{HYPHENS_CLASS}\s]{{6,}})\d"
)

def pick_phone_token_raw(line: str) -> str:
    """1行から電話番号らしい文字列を抽出（9〜11桁のみ）"""
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
        if not digits.startswith("0") and not digits.startswith("81"):
            continue
        score = (len(digits), tok.count("-"))
        cands.append((score, tok))
    if not cands:
        return ""
    cands.sort(key=lambda x: x[0], reverse=True)
    return cands[0][1]

def phone_digits_only(s: str) -> str:
    """内部照合用に数字だけ抽出"""
    return re.sub(r"\D", "", str(s or ""))

# ===============================
# Google検索リスト形式
# ===============================
def extract_google_vertical(lines):
    results = []
    rows = [str(l) for l in lines if str(l).strip() != ""]
    for i, line in enumerate(rows):
        ph_raw = pick_phone_token_raw(line)
        if ph_raw:
            phone = ph_raw  # 原文保持
            address = rows[i - 1] if i - 1 >= 0 else ""
            industry = extract_industry(rows[i - 2]) if i - 2 >= 0 else ""
            company = rows[i - 3] if i - 3 >= 0 else ""
            results.append([company, industry, clean_address(address), phone])
    return pd.DataFrame(results, columns=["企業名", "業種", "住所", "電話番号"])

# ===============================
# シゴトアルワ形式
# ===============================
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
        if k in ["住所", "所在地"]:
            current["住所"] = clean_address(v)
        elif k in ["電話", "電話番号", "TEL", "Tel"]:
            current["電話番号"] = v
        elif k in ["業種", "事業内容"]:
            current["業種"] = extract_industry(v)
        elif k and not v:
            if current["企業名"]:
                flush()
            current["企業名"] = k
    if current["企業名"]:
        flush()
    return pd.DataFrame(out, columns=["企業名", "業種", "住所", "電話番号"])

# ===============================
# 日本倉庫協会形式
# ===============================
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
# 共通整形
# ===============================
def clean_dataframe_except_phone(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ["企業名", "業種", "住所"]:
        df[c] = df[c].map(normalize_text)
    return df.fillna("")

# ===============================
# アップロード処理
# ===============================
st.markdown("### 📤 整形対象のExcelファイルをアップロード")
profile = st.selectbox(
    "抽出プロファイル",
    ["Google検索リスト（縦読み・電話上下型）", "シゴトアルワ検索リスト（縦積み）", "日本倉庫協会リスト（4列型）"]
)
uploaded_file = st.file_uploader("Excelファイルを選択", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    if profile == "Google検索リスト（縦読み・電話上下型）":
        df0 = pd.read_excel(uploaded_file, header=None).fillna("")
        lines = df0.iloc[:, 0].tolist()
        df = extract_google_vertical(lines)
    elif profile == "シゴトアルワ検索リスト（縦積み）":
        df0 = pd.read_excel(xl, header=None).fillna("")
        df = extract_shigoto_arua(df0)
    else:
        df0 = pd.read_excel(xl, header=None).fillna("")
        df = extract_warehouse_association(df0)

    df = clean_dataframe_except_phone(df)
    df["__digits"] = df["電話番号"].map(phone_digits_only)

    st.success(f"✅ 整形完了：{len(df)}件の企業データを取得しました。")
    edited = st.data_editor(df[["企業名", "業種", "住所", "電話番号"]], use_container_width=True)

    if st.button("✅ この内容で確定（反映）"):
        df["企業名"], df["業種"], df["住所"], df["電話番号"] = (
            edited["企業名"],
            edited["業種"],
            edited["住所"],
            edited["電話番号"],
        )
        st.success("編集内容を反映しました。出力はこの表記のままです。")

    # Excel出力
    out = io.BytesIO()
    df_out = df.drop(columns=["__digits"], errors="ignore")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="出力")
    out.seek(0)
    st.download_button("📥 整形済みリストをダウンロード", data=out, file_name="整形済みリスト.xlsx")

else:
    st.info("Excelファイルをアップロードしてください。")
