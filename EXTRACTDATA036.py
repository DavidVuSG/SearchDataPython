import pandas as pd

# =========================
# COMMON CLEAN FUNCTION
# =========================
def clean_file(path):
    df = pd.read_excel(path, header=None)

    # Remove first 3 rows, row 4 = header
    df = df.iloc[3:].reset_index(drop=True)

    df.columns = [
        "NO", "LOC", "MAHANG", "TENHANG", "SOLUONG",
        "SOTHUNG", "LPN", "PO", "SUPPLIER",
        "DATEREC", "DATEPOST"
    ]

    # =========================
    # REMOVE HEADER ROWS INSIDE DATA
    # =========================
    df = df[
        df["MAHANG"].notna() &
        (df["MAHANG"] != "") &
        (df["MAHANG"] != "MÃ HÀNG")
    ]

    # =========================
    # SAFE NUMERIC CONVERSION
    # =========================
    df["SOLUONG"] = pd.to_numeric(df["SOLUONG"], errors="coerce")
    df["SOTHUNG"] = pd.to_numeric(df["SOTHUNG"], errors="coerce")

    df = df.dropna(subset=["SOLUONG"])

    # =========================
    # DATE PARSE (FIFO)
    # =========================
    df["DATEREC"] = pd.to_datetime(
        df["DATEREC"],
        format="%d/%m/%Y",
        errors="coerce"
    )

    # Remove rows without date
    df = df[df["DATEREC"].notna()]

    df["LOT6"] = df["DATEREC"].dt.strftime("%Y%m%d")

    # =========================
    # SPLIT MAHANG
    # =========================
    df["ID"] = df["MAHANG"].astype(str).str.extract(r"^[^_]+_([^_]+)_")
    df["VER"] = df["MAHANG"].astype(str).str.extract(r"(_\d+)$")

    return df


# =========================
# 1. FILE b37.xls
# =========================
df_037 = clean_file("b37.xls")

hanghu_keys = ["A2-04", "STAGE", "INTRANSIT"]
df_037["STATUS"] = df_037["LOC"].astype(str).apply(
    lambda x: "HANGHU" if any(k in x for k in hanghu_keys) else "HANGHOLD"
)

d_hanghu = df_037[df_037["STATUS"] == "HANGHU"].copy()
d_hanghold = df_037[df_037["STATUS"] == "HANGHOLD"].copy()

# FIFO
d_hanghu.sort_values(["MAHANG", "DATEREC", "LPN"], inplace=True)
d_hanghold.sort_values(["MAHANG", "DATEREC", "LPN"], inplace=True)


# =========================
# 2. FILE b36.xls
# =========================
df_036 = clean_file("b36.xls")

# Drop LOC contains PICKTO, AGV
df_036 = df_036[
    ~df_036["LOC"].astype(str).str.contains("PICKTO|AGV", regex=True, na=False)
]

df_036["STATUS"] = "HANGOK"
d_hangok = df_036.copy()


# =========================
# MERGE HANGBTP
# =========================
d_hangbtp = pd.concat([d_hanghold, d_hangok], ignore_index=True)

# FIFO FINAL
d_hangbtp.sort_values(["MAHANG", "DATEREC", "LPN"], inplace=True)


# =========================
# EXPORT
# =========================
d_hanghu.to_excel("HANGHU.xlsx", index=False)
d_hangbtp.to_excel("HANGBTP.xlsx", index=False)

 