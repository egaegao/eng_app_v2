import streamlit as st
import pandas as pd
from io import BytesIO
from dashboard_config import REQUIRED_SHEETS

# Impor fungsi analisis dari backend
try:
    from weekly_analysis import run_analysis_from_data
except ImportError:
    # Fallback jika struktur folder berbeda saat runtime
    try:
        from scripts.weekly_analysis import run_analysis_from_data
    except ImportError:
        st.error("Gagal mengimpor 'run_analysis_from_data'. Pastikan file 'weekly_analysis.py' berada di folder 'scripts'.")
        st.stop()

def validate_workbook(uploaded_file):
    """
    Validasi file Excel untuk memastikan semua sheet yang dibutuhkan tersedia.
    """
    try:
        # Menggunakan pd.ExcelFile tanpa membaca seluruh data untuk efisiensi validasi
        excel_file = pd.ExcelFile(uploaded_file)
        missing_sheets = [s for s in REQUIRED_SHEETS if s not in excel_file.sheet_names]
        return excel_file, missing_sheets
    except Exception as e:
        st.error(f"Error saat validasi file: {e}")
        return None, REQUIRED_SHEETS

@st.cache_data(show_spinner="Membaca Data Excel...")
def load_data_from_workbook(file_bytes: bytes):
    """
    Membaca semua sheet dari workbook Excel ke dalam DataFrames.
    Menggunakan st.cache_data agar proses baca file tidak diulang-ulang.
    """
    workbook = BytesIO(file_bytes)
    
    # Membaca tiap sheet sesuai kebutuhan dashboard
    weekly_kpi = pd.read_excel(workbook, sheet_name="weekly_kpi")
    fuel_ratio = pd.read_excel(workbook, sheet_name="fuel_ratio")
    issue_log = pd.read_excel(workbook, sheet_name="issue_log")
    action_tracker = pd.read_excel(workbook, sheet_name="action_tracker")
    inventory = pd.read_excel(workbook, sheet_name="inventory")
    hauling_review = pd.read_excel(workbook, sheet_name="hauling_review")
    findings_summary = pd.read_excel(workbook, sheet_name="findings_summary")
    unit_performance = pd.read_excel(workbook, sheet_name="unit_performance") 

    return (
        weekly_kpi, fuel_ratio, issue_log, action_tracker, 
        inventory, hauling_review, findings_summary,
        unit_performance 
    )

@st.cache_data(show_spinner="Menjalankan Analisis Performa...")
def run_analysis_cached(file_hash, wk, fr, il, at, inv, hr, fs):
    """
    Menjalankan mesin analisis backend. 
    Menggunakan file_hash sebagai key agar analisis hanya diulang jika file berubah.
    Fungsi ini hanya menerima 7 dataset (unit_performance tidak diproses di backend analisis).
    Menggunakan .copy() untuk menjaga integritas data asli.
    """
    return run_analysis_from_data(
        wk.copy(), 
        fr.copy(), 
        il.copy(), 
        at.copy(),
        inv.copy(), 
        hr.copy(), 
        fs.copy()
    )