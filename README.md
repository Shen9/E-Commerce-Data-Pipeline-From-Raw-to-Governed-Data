# E-Commerce Data Pipeline: From Raw to Governed Data

Proyek ini merupakan implementasi *end-to-end data pipeline* sederhana menggunakan Python untuk mengolah data mentah e-commerce menjadi data yang bersih, tervalidasi, dan memenuhi standar tata kelola data (*Data Governance*). 

Proyek ini dikembangkan sebagai bagian dari sertifikasi **Data Engineer di Dibimbing.id**.

## 🚀 Fitur Utama

1.  **Automated Data Loading**: Membaca data mentah dari format Excel dengan penanganan tipe data otomatis.
2.  **Data Profiling**: Analisis cepat terhadap keunikan data dan persentase *missing values*.
3.  **Data Quality (DQ) Check**: Validasi ketat terhadap kolom kritis, deteksi duplikat, pengecekan tipe data, hingga deteksi *outlier* menggunakan metode IQR.
4.  **Robust Data Cleansing**: Pembersihan otomatis baris bermasalah, standarisasi status order, dan kalkulasi ulang nilai transaksi untuk memastikan integritas data.
5.  **PII Masking (Data Governance)**: Implementasi teknik *masking* pada data sensitif (Nama, Email, Telepon) yang selaras dengan **UU PDP No. 27 Tahun 2022** dan standar **ISO 27701**.
6.  **Automated Reporting**: Menghasilkan output Excel dengan tiga sheet: Data Bersih, Laporan Kualitas Data (DQ Report), dan Ringkasan KPI.

## 🛠️ Teknologi yang Digunakan

-   **Bahasa Pemrograman**: Python 3.x
-   **Library Utama**: `pandas`, `numpy`, `openpyxl`
-   **Package Manager**: `uv` (modern & fast Python package manager)

## 📁 Struktur Proyek

-   `ecommercePipeline.py`: Script utama yang berisi logika ETL dan Governance.
-   `ecommerce_data_engineering_project_raw.xlsx`: Dataset mentah (input).
-   `ecommerce_clean_output.xlsx`: Hasil akhir pemrosesan (output).
-   `pyproject.toml` & `uv.lock`: Konfigurasi dependensi proyek.

## ⚙️ Cara Menjalankan

1.  Pastikan Anda telah menginstal `uv` atau `pip`.
2.  Instal dependensi yang diperlukan:
    ```bash
    uv add pandas openpyxl
    ```
3.  Jalankan pipeline:
    ```bash
    python ecommercePipeline.py
    ```

## 📊 Alur Kerja Pipeline (Data Flow)

1.  **Load**: Membaca sheet `RAW_DATA`.
2.  **Profiling**: Mengecek kesehatan data awal.
3.  **Quality Check**: Mengidentifikasi baris yang akan di-reject (null, duplikat, format salah).
4.  **Cleansing**: Eksekusi pembersihan dan standarisasi.
5.  **Masking**: Melindungi privasi pelanggan dengan menyamarkan identitas (PII).
6.  **Export**: Menyimpan hasil ke file Excel baru.
7.  **Summary**: Menampilkan statistik akhir durasi dan *pass rate* data.

## 🛡️ Standar Masking (Governance)

| Kolom Asli | Hasil Masking (Contoh) |
| :--- | :--- |
| customer_name | A*** P*** |
| customer_email | an****@gmail.com |
| phone_number | 0812-****-7890 |

---
*Dibuat untuk portofolio Data Engineering.*