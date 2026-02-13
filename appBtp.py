import streamlit as st
import pandas as pd
from pathlib import Path

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="BTP Inventory Search", layout="wide")

B36 = "b36.xls"
B37 = "b37.xls"
OUT_BTP = "HANGBTP.xlsx"
OUT_HU = "HANGHU.xlsx"

# =========================
# COMMON CLEAN FUNCTION
# =========================
def clean_file(path):
    df = pd.read_excel(path, header=None)

    # bỏ 3 dòng đầu, dòng 4 là header
    df = df.iloc[3:].reset_index(drop=True)

    df.columns = [
        "NO", "LOC", "MAHANG", "TENHANG", "SOLUONG",
        "SOTHUNG", "LPN", "PO", "SUPPLIER",
        "DATEREC", "DATEPOST"
    ]

    # bỏ dòng rác / header trôi
    df = df[
        df["MAHANG"].notna()
        & (df["MAHANG"] != "")
        & (df["MAHANG"] != "MÃ HÀNG")
    ]

    # ép số
    df["SOLUONG"] = pd.to_numeric(df["SOLUONG"], errors="coerce")
    df["SOTHUNG"] = pd.to_numeric(df["SOTHUNG"], errors="coerce")

    # ép ngày DD/MM/YYYY
    df["DATEREC"] = pd.to_datetime(
        df["DATEREC"],
        dayfirst=True,
        errors="coerce"
    )

    return df


# =========================
# EXTRACT PIPELINE
# =========================
def extract_data():
    # ---------- b37 → HANGHOLD / HANGHU
    df_037 = clean_file(B37)

    hanghu_keys = ["A2-04", "STAGE", "INTRANSIT"]
    df_037["STATUS"] = df_037["LOC"].astype(str).apply(
        lambda x: "HANGHU" if any(k in x for k in hanghu_keys) else "HANGHOLD"
    )

    d_hanghu = df_037[df_037["STATUS"] == "HANGHU"].copy()
    d_hanghold = df_037[df_037["STATUS"] == "HANGHOLD"].copy()

    # ---------- b36 → HANGOK
    df_036 = clean_file(B36)
    df_036 = df_036[
        ~df_036["LOC"].astype(str).str.contains("PICKTO|AGV", regex=True)
    ]
    df_036["STATUS"] = "HANGOK"
    d_hangok = df_036.copy()

    # ---------- MERGE BTP
    d_hangbtp = pd.concat([d_hanghold, d_hangok], ignore_index=True)

    # FIFO
    d_hangbtp = d_hangbtp[
        d_hangbtp["DATEREC"].notna()
    ].sort_values("DATEREC")

    d_hanghu = d_hanghu[
        d_hanghu["DATEREC"].notna()
    ].sort_values("DATEREC")

    # ---------- EXPORT
    d_hangbtp.to_excel(OUT_BTP, index=False)
    d_hanghu.to_excel(OUT_HU, index=False)


# =========================
# LOAD DATA (CACHE)
# =========================
@st.cache_data
def load_data():
    df1 = pd.read_excel(OUT_BTP)
    df2 = pd.read_excel(OUT_HU)
    return pd.concat([df1, df2], ignore_index=True)


# =========================
# UI
# =========================
st.title("📦 BTP Inventory Interactive Search")

if st.button("🔄 Reload data"):
    extract_data()
    st.cache_data.clear()
    st.rerun()

if not Path(OUT_BTP).exists():
    extract_data()

df = load_data()


# =========================
# LOAD PACK FILE
# =========================
PACK_FILE = "PACK PPL MPE IMPORT.xlsx"

if Path(PACK_FILE).exists():
    df_pack = pd.read_excel(PACK_FILE, usecols=[0, 2])
    df_pack.columns = ["MAHANG", "PACK"]

    # ép PACK về số
    df_pack["PACK"] = pd.to_numeric(df_pack["PACK"], errors="coerce")

    # merge
    df = df.merge(df_pack, on="MAHANG", how="left")
else:
    df["PACK"] = None



# =========================
# SEARCH INPUT
# =========================
status = st.selectbox(
    "STATUS",
    ["ALL", "HANGOK", "HANGHOLD", "HANGHU"]
)

search_text = st.text_input(
    "🔍 Search (MAHANG, PO, LOC, LPN,SUPPLIER) – dùng dấu ,",
    placeholder="VD: RPL-9T, 1169, A1-02-03-1"
)



# =========================
# SEARCH LOGIC
# =========================
if search_text:
    tokens = [t.strip() for t in search_text.split(",")]
    cols = ["MAHANG", "PO", "LOC", "LPN","SUPPLIER"]

    for col, token in zip(cols, tokens):
        if token:
            df = df[df[col].astype(str).str.contains(token, case=False, na=False)]

if status != "ALL":
    df = df[df["STATUS"] == status]

# FIFO lần cuối
df["DATEREC"] = pd.to_datetime(
    df["DATEREC"],
    dayfirst=True,
    errors="coerce"
)

df = df[df["DATEREC"].notna()]
df = df.sort_values("DATEREC")

# =========================
# RESULT (CHỈ SHOW CỘT CẦN)
# =========================
result = df[
    ["LOC", "LPN", "MAHANG", "SOLUONG", "SOTHUNG", "PO", "SUPPLIER","DATEREC","PACK", "STATUS"]
]

st.markdown(f"### 📄 Result: **{len(result)} rows**")

st.dataframe(result, width="stretch")

# =========================
# EXPORT
# =========================
from io import BytesIO

output = BytesIO()
result.to_excel(output, index=False)
output.seek(0)

st.download_button(
    label="⬇ Export result to Excel",
    data=output,
    file_name="result.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
