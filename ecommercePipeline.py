import pandas as pd
import numpy as np
import os
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

FILE_PATH   = os.path.join(os.path.dirname(__file__), "ecommerce_data_engineering_project_raw.xlsx")
SHEET_RAW   = "RAW_DATA"
OUTPUT_PATH = os.path.dirname(FILE_PATH)

VALID_STATUSES = {"Completed", "Pending", "Cancelled", "Shipped", "Returned"}
STATUS_MAP     = {"DONE": "Completed", "process": "Pending", "cancel": "Cancelled", "N/A": "Pending"}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER
# ─────────────────────────────────────────────────────────────────────────────

def separator(title=""):
    print("\n" + "=" * 65)
    if title:
        print(f"  {title}")
        print("=" * 65)

def log(msg):
    print(f"  >> {msg}")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────

def load_data(path, sheet):
    separator("STEP 1 — LOAD RAW DATA")
    log(f"File  : {path}")
    log(f"Sheet : {sheet}")

    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = df.columns.str.strip().str.lower()

    log(f"Berhasil dimuat : {df.shape[0]:,} baris, {df.shape[1]} kolom")
    log(f"Kolom           : {list(df.columns)}")
    return df

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — DATA PROFILING
# ─────────────────────────────────────────────────────────────────────────────

def profiling(df):
    separator("STEP 2 — DATA PROFILING")

    log("Tipe data tiap kolom:")
    for col in df.columns:
        null_count = df[col].isnull().sum()
        null_pct   = null_count / len(df) * 100
        unique     = df[col].nunique()
        print(f"     {col:<25} | unique: {unique:>5} | null: {null_count:>4} ({null_pct:.1f}%)")

    print()
    log(f"Total baris : {len(df):,}")
    log(f"Total kolom : {len(df.columns)}")

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — DATA QUALITY CHECK
# ─────────────────────────────────────────────────────────────────────────────

def data_quality_check(df):
    separator("STEP 3 — DATA QUALITY CHECK")
    issues = {}

    # 3a. Missing values pada kolom kritis
    critical_cols = ["order_id", "customer_id", "order_date", "product_id",
                     "quantity", "unit_price", "total_price", "order_status"]
    for col in critical_cols:
        if col in df.columns:
            n = df[col].isnull().sum()
            if n > 0:
                issues[f"null_{col}"] = n
                log(f"[NULL]      {col}: {n} baris kosong")

    # 3b. Duplikat order_id
    dup = df.duplicated(subset=["order_id"], keep=False).sum()
    if dup > 0:
        issues["duplicate_order_id"] = dup
        log(f"[DUPLIKAT]  order_id duplikat: {dup} baris")

    # 3c. Quantity tidak valid (harus angka > 0)
    df["_qty_num"] = pd.to_numeric(df["quantity"], errors="coerce")
    bad_qty = df["_qty_num"].isnull() | (df["_qty_num"] <= 0)
    n = bad_qty.sum()
    if n > 0:
        issues["invalid_quantity"] = n
        log(f"[INVALID]   quantity <= 0 atau bukan angka: {n} baris")

    # 3d. Total price tidak valid
    df["_total_num"] = pd.to_numeric(df["total_price"], errors="coerce")
    bad_total = df["_total_num"].isnull() | (df["_total_num"] < 0)
    n = bad_total.sum()
    if n > 0:
        issues["invalid_total_price"] = n
        log(f"[INVALID]   total_price negatif atau bukan angka: {n} baris")

    # 3e. Format tanggal tidak valid
    df["_date_parsed"] = pd.to_datetime(df["order_date"], format="%Y-%m-%d", errors="coerce")
    bad_date = df["_date_parsed"].isnull()
    n = bad_date.sum()
    if n > 0:
        issues["invalid_order_date"] = n
        log(f"[INVALID]   order_date format salah: {n} baris")

    # 3f. Status tidak sesuai standar
    all_known = VALID_STATUSES | set(STATUS_MAP.keys())
    bad_status = ~df["order_status"].isin(all_known) & df["order_status"].notna()
    n = bad_status.sum()
    if n > 0:
        issues["nonstandard_status"] = n
        log(f"[NONSTANDARD] order_status tidak dikenal: {n} baris")

    # 3g. Outlier total_price (IQR method)
    valid_prices = df["_total_num"].dropna()
    Q1, Q3  = valid_prices.quantile(0.25), valid_prices.quantile(0.75)
    IQR     = Q3 - Q1
    lower   = Q1 - 1.5 * IQR
    upper   = Q3 + 1.5 * IQR
    outlier = ((df["_total_num"] < lower) | (df["_total_num"] > upper)).sum()
    log(f"[OUTLIER]   total_price di luar IQR range [{lower:,.0f} – {upper:,.0f}]: {outlier} baris (info saja, tidak di-reject)")

    total_issues = sum(issues.values())
    log(f"\n  Total baris bermasalah (akan di-reject): {total_issues}")

    return df, issues

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — DATA CLEANSING
# ─────────────────────────────────────────────────────────────────────────────

def cleansing(df):
    separator("STEP 4 — DATA CLEANSING")
    before = len(df)

    # Hapus baris dengan kolom kritis null
    critical_cols = ["order_id", "order_date", "product_id", "quantity",
                     "unit_price", "total_price", "order_status"]
    df = df.dropna(subset=[c for c in critical_cols if c in df.columns])
    log(f"Setelah hapus null kritis     : {len(df):,} baris (hapus {before - len(df)})")

    # Hapus duplikat order_id — keep first
    before_dup = len(df)
    df = df.drop_duplicates(subset=["order_id"], keep="first")
    log(f"Setelah hapus duplikat        : {len(df):,} baris (hapus {before_dup - len(df)})")

    # Konversi tipe data
    df["quantity"]   = pd.to_numeric(df["quantity"],   errors="coerce")
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")
    df["total_price"]= pd.to_numeric(df["total_price"],errors="coerce")
    df["discount_pct"]= pd.to_numeric(df["discount_pct"], errors="coerce").fillna(0)
    df["order_date"] = pd.to_datetime(df["order_date"], format="%Y-%m-%d", errors="coerce")

    # Hapus baris dengan nilai tidak valid setelah konversi
    before_inv = len(df)
    df = df.dropna(subset=["quantity", "unit_price", "total_price", "order_date"])
    df = df[df["quantity"] > 0]
    df = df[df["total_price"] >= 0]
    log(f"Setelah validasi tipe & range : {len(df):,} baris (hapus {before_inv - len(df)})")

    # Standardisasi order_status
    before_st = len(df)
    df["order_status"] = df["order_status"].map(
        lambda x: STATUS_MAP.get(x, x) if x not in VALID_STATUSES else x
    )
    df = df[df["order_status"].isin(VALID_STATUSES)]
    log(f"Setelah standarisasi status   : {len(df):,} baris (hapus {before_st - len(df)})")

    # Recalculate total_price (pastikan konsisten)
    df["total_price"] = (df["quantity"] * df["unit_price"] * (1 - df["discount_pct"] / 100)).round(0)

    # Bersihkan kolom temporary
    df = df[[c for c in df.columns if not c.startswith("_")]]

    rejected = before - len(df)
    pass_rate = len(df) / before * 100
    log(f"\n  Raw        : {before:,} baris")
    log(f"  Clean      : {len(df):,} baris")
    log(f"  Rejected   : {rejected:,} baris")
    log(f"  Pass rate  : {pass_rate:.1f}%")

    return df

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — PII MASKING (Data Governance)
# ─────────────────────────────────────────────────────────────────────────────

def mask_name(name):
    if pd.isnull(name) or str(name).strip() == "":
        return "UNKNOWN"
    parts = str(name).split()
    return " ".join((p[0] + "***") if len(p) > 1 else p for p in parts)

def mask_email(email):
    if pd.isnull(email) or "@" not in str(email):
        return "UNKNOWN"
    user, domain = str(email).split("@", 1)
    return user[:2] + "****@" + domain

def mask_phone(phone):
    if pd.isnull(phone) or str(phone).strip() == "":
        return "UNKNOWN"
    p = str(phone)
    return p[:4] + "****" + p[-4:] if len(p) >= 8 else "****"

def apply_pii_masking(df):
    separator("STEP 5 — PII MASKING (Data Governance)")

    df = df.copy()
    df["customer_name_masked"]  = df["customer_name"].apply(mask_name)
    df["customer_email_masked"] = df["customer_email"].apply(mask_email)
    df["phone_masked"]          = df["phone_number"].apply(mask_phone)

    # Hapus kolom PII asli dari output
    df = df.drop(columns=["customer_name", "customer_email", "phone_number"], errors="ignore")

    log("Kolom PII asli dihapus dari output clean")
    log("Masking diterapkan pada:")
    log("  customer_name  → customer_name_masked  (contoh: Andi Pratama → A*** P***)")
    log("  customer_email → customer_email_masked  (contoh: andi@gmail.com → an****@gmail.com)")
    log("  phone_number   → phone_masked           (contoh: 0812-3456-7890 → 0812-****-7890)")
    log("Standar referensi: ISO 27701 / UU PDP Indonesia No. 27 Tahun 2022")

    return df

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — EXPORT KE EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def export_to_excel(df_clean, issues, total_raw, output_path):
    separator("STEP 6 — EXPORT KE EXCEL")

    out_file = os.path.join(output_path, "ecommerce_clean_output.xlsx")

    try:
        with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
            # Sheet 1: Clean Data
            df_clean.to_excel(writer, sheet_name="CLEANED_DATA", index=False)
            log(f"Sheet CLEANED_DATA    : {len(df_clean):,} baris")

            # Sheet 2: Data Quality Report
            report_rows = []
            for issue, count in issues.items():
                report_rows.append({"issue_type": issue, "row_count": count,
                                     "action": "Rows removed"})
            report_rows.append({
                "issue_type": "TOTAL_RAW",         "row_count": total_raw,          "action": "-"})
            report_rows.append({
                "issue_type": "TOTAL_CLEAN",        "row_count": len(df_clean),      "action": "-"})
            report_rows.append({
                "issue_type": "PASS_RATE_%",        "row_count": round(len(df_clean)/total_raw*100, 1), "action": "-"})

            pd.DataFrame(report_rows).to_excel(writer, sheet_name="DQ_REPORT", index=False)
            log(f"Sheet DQ_REPORT       : {len(report_rows)} baris")

            # Sheet 3: Summary Stats
            stats = {
                "Total Orders (clean)":     len(df_clean),
                "Total Revenue (IDR)":      int(df_clean["total_price"].sum()),
                "Avg Order Value (IDR)":    int(df_clean["total_price"].mean()),
                "Max Order Value (IDR)":    int(df_clean["total_price"].max()),
                "Min Order Value (IDR)":    int(df_clean["total_price"].min()),
                "Total Units Sold":         int(df_clean["quantity"].sum()),
                "Avg Discount (%)":         round(df_clean["discount_pct"].mean(), 1),
            }
            pd.DataFrame(stats.items(), columns=["Metric", "Value"]).to_excel(
                writer, sheet_name="SUMMARY_STATS", index=False)
            log(f"Sheet SUMMARY_STATS   : {len(stats)} KPI")
    except ImportError:
        log("ERROR: Library 'openpyxl' tidak ditemukan. Pipeline gagal mengekspor ke Excel.")
        log("Saran: Jalankan perintah 'uv add openpyxl' untuk menambahkannya ke proyek.")
        return None

    log(f"\n  File tersimpan di: {out_file}")
    return out_file

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7 — FINAL SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def final_summary(df_clean, total_raw, start_time):
    separator("STEP 7 — FINAL SUMMARY")

    duration = (datetime.now() - start_time).total_seconds()

    log(f"Total raw records   : {total_raw:,}")
    log(f"Total clean records : {len(df_clean):,}")
    log(f"Rejected records    : {total_raw - len(df_clean):,}")
    log(f"Pass rate           : {len(df_clean)/total_raw*100:.1f}%")
    log(f"Durasi proses       : {duration:.2f} detik")
    log(f"Selesai pada        : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    print()
    log("Pipeline selesai! File output siap digunakan.")
    separator()

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — JALANKAN SEMUA STEP
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    start = datetime.now()

    print()
    print("=" * 65)
    print("  E-COMMERCE DATA PIPELINE — FROM RAW TO GOVERNED DATA")
    print("  Dibimbing.id | Data Engineer Certification Project")
    print("=" * 65)

    # Jalankan pipeline
    df_raw              = load_data(FILE_PATH, SHEET_RAW)
    total_raw           = len(df_raw)

    profiling(df_raw)

    df_checked, issues  = data_quality_check(df_raw)
    df_clean            = cleansing(df_checked)
    df_masked           = apply_pii_masking(df_clean)

    export_to_excel(df_masked, issues, total_raw, OUTPUT_PATH)
    final_summary(df_masked, total_raw, start)