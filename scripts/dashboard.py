import streamlit as st
import pandas as pd
import hashlib
from io import BytesIO # Diperlukan untuk proses download Excel

# Import modul login yang baru dibuat
from dashboard_login import login_screen

from dashboard_styles import apply_custom_styles
from dashboard_helpers import parse_week_date, normalize_text, get_period_dates
from dashboard_loader import validate_workbook, load_data_from_workbook, run_analysis_cached
from dashboard_filters import filter_by_block, filter_exact_week, filter_range

from dashboard_pages.weekly_page import render_weekly_page
from dashboard_pages.range_page import render_range_page

# ==========================================
# 1. PAGE CONFIG & LOGIN
# ==========================================
st.set_page_config(
    page_title="Mining Weekly Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Jalankan sistem login di awal
login_screen()

# Jika lolos login (authenticated == True), baru jalankan sisa script di bawah ini
apply_custom_styles()

# --- FUNGSI HELPER: KONVERSI DATAFRAME KE EXCEL PROFESIONAL ---
def to_excel(df):
    output = BytesIO()
    # Menggunakan engine xlsxwriter untuk fitur styling rapi
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Mining_Report')
        workbook = writer.book
        worksheet = writer.sheets['Mining_Report']
        
        # Format Header (Tebal & Background Abu-abu)
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#F2F2F2', 
            'border': 1
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            # Logika Auto-adjust lebar kolom agar tidak terpotong
            if not df.empty:
                column_len = max(df[value].astype(str).str.len().max(), len(value)) + 2
            else:
                column_len = len(value) + 2
            worksheet.set_column(col_num, col_num, column_len)
            
    return output.getvalue()

# ==========================================
# 2. FILE UPLOAD & VALIDATION
# ==========================================
# Tombol Logout diletakkan di paling atas sidebar setelah login
if st.sidebar.button("🔒 Log Out"):
    st.session_state['authenticated'] = False
    st.rerun()

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel Master File",
    type=["xlsx"],
    help="Upload 1 workbook Excel yang berisi semua data per sheet."
)

if uploaded_file is None:
    st.title("📊 Weekly Mining Dashboard")
    st.info("Selamat Datang. Silakan upload file Excel master terlebih dahulu di sidebar.")
    st.caption("Professional dashboard generated from integrated weekly mining data.")
    st.stop()

# --- TAG INFO: LOAD SECEPAT KILAT ---
st.sidebar.markdown(
    """
    <div style="font-size: 0.85rem; font-weight: 700; color: #475569; text-align: left; margin-bottom: 20px;">
        Learning App | Indrawijaya, CEE | 2026
    </div>
    <hr style="margin-top: -10px; margin-bottom: 20px; border-top: 1px solid #e2e8f0;">
    """, 
    unsafe_allow_html=True
)

excel_file, missing_sheets = validate_workbook(uploaded_file)

if excel_file is None:
    st.error("Gagal membaca file Excel. File mungkin rusak atau format tidak didukung.")
    st.stop()

if missing_sheets:
    st.error(f"Sheet berikut belum ada di file: {', '.join(missing_sheets)}")
    st.stop()

# ==========================================
# 3. LOAD DATA & CACHE ANALYSIS
# ==========================================
file_bytes = uploaded_file.getvalue()
(
    weekly_kpi, fuel_ratio, issue_log, action_tracker, 
    inventory, hauling_review, findings_summary,
    unit_performance
) = load_data_from_workbook(file_bytes)

file_hash = hashlib.md5(file_bytes).hexdigest()

# REVISI: unit_performance dihapus dari argumen run_analysis_cached sesuai instruksi
analysis_result = run_analysis_cached(
    file_hash, weekly_kpi, fuel_ratio, issue_log, action_tracker, 
    inventory, hauling_review, findings_summary
)

# Sinkronisasi Dataframe Utama dengan Hasil Analisis
if "kpi_summary" in analysis_result:
    weekly_kpi = analysis_result["kpi_summary"]
if "fuel_ratio" in analysis_result:
    fuel_ratio = analysis_result["fuel_ratio"]
if "hauling_review" in analysis_result:
    hauling_review = analysis_result["hauling_review"]

# ==========================================
# 4. CLEANING & PROCESSING
# ==========================================
required_cols = ["metric", "plan", "actual", "week_date", "block"]
missing_cols = [c for c in required_cols if c not in weekly_kpi.columns]
if missing_cols:
    st.error(f"Kolom wajib di 'weekly_kpi' tidak ditemukan: {missing_cols}")
    st.stop()

weekly_kpi = parse_week_date(weekly_kpi)
issue_log = parse_week_date(issue_log)
action_tracker = parse_week_date(action_tracker)
inventory = parse_week_date(inventory)
findings_summary = parse_week_date(findings_summary)
unit_performance = parse_week_date(unit_performance)

if "due_date" in action_tracker.columns:
    action_tracker["due_date"] = pd.to_datetime(
        action_tracker["due_date"], format="mixed", dayfirst=True, errors="coerce"
    )

weekly_kpi = weekly_kpi.dropna(subset=["week_date"]).copy()
if weekly_kpi.empty:
    st.error("Data 'weekly_kpi' kosong atau tidak ada tanggal (week_date) yang valid.")
    st.stop()

if "severity" in issue_log.columns:
    issue_log["severity"] = issue_log["severity"].astype(str).str.strip().str.title()
if "status" in issue_log.columns:
    issue_log["status"] = normalize_text(issue_log["status"])
if "status" in action_tracker.columns:
    action_tracker["status"] = normalize_text(action_tracker["status"])
if "priority" in action_tracker.columns:
    action_tracker["priority"] = normalize_text(action_tracker["priority"])
if "inventory_type" in inventory.columns:
    inventory["inventory_type"] = normalize_text(inventory["inventory_type"])
if "activity" in fuel_ratio.columns:
    fuel_ratio["activity"] = normalize_text(fuel_ratio["activity"])
if "category" in findings_summary.columns:
    findings_summary["category"] = normalize_text(findings_summary["category"])

# Overdue logic
if "due_date" in action_tracker.columns and "week_date" in action_tracker.columns:
    action_tracker["overdue_flag"] = (
        (action_tracker["due_date"].notna()) &
        (action_tracker["week_date"].notna()) &
        (action_tracker["status"].isin(["Open", "Progress"])) &
        (action_tracker["due_date"] < action_tracker["week_date"])
    )

# ==========================================
# 5. SIDEBAR FILTERS
# ==========================================
st.sidebar.title("Dashboard Filter")
page_mode = st.sidebar.radio("Select Page", ["Weekly", "Range Analysis"])

block_list = sorted(weekly_kpi["block"].dropna().astype(str).unique())
if not block_list:
    st.error("Tidak ada data block yang valid ditemukan.")
    st.stop()

block = st.sidebar.selectbox("Select Mining Block", block_list)

block_df = filter_by_block(weekly_kpi, block)
block_issue_log = filter_by_block(issue_log, block)
block_action_tracker = filter_by_block(action_tracker, block)
block_inventory = filter_by_block(inventory, block)
block_fuel_ratio = filter_by_block(fuel_ratio, block)
block_findings = filter_by_block(findings_summary, block)
block_hauling_review = filter_by_block(hauling_review, block)
block_unit_performance = filter_by_block(unit_performance, block)

# ==========================================
# 6. PAGE ROUTING
# ==========================================
if page_mode == "Weekly":
    available_weeks = sorted(block_df["week_date"].dropna().unique())
    if not available_weeks:
        st.warning(f"Tidak ada data mingguan tersedia untuk block {block}.")
        st.stop()

    selected_week = st.sidebar.selectbox(
        "Select Week",
        available_weeks,
        index=len(available_weeks) - 1,
        format_func=lambda x: pd.Timestamp(x).strftime("%d %b %Y")
    )

    latest_df = block_df[block_df["week_date"] == selected_week].copy()
    previous_weeks = [w for w in available_weeks if w < selected_week]
    prev_df = block_df[block_df["week_date"] == previous_weeks[-1]].copy() if previous_weeks else None

    # --- FIX 1: OVERRIDE BLOCK UNIT PERFORMANCE ---
    # Mengambil data unit performance untuk seluruh block pada minggu yang dipilih
    weekly_unit_full = filter_exact_week(unit_performance, selected_week)

    # Routing ke render fungsi Weekly Page
    render_weekly_page(
        block=block, 
        selected_week=selected_week, 
        available_weeks=available_weeks,
        block_df=block_df, 
        weekly_issue=filter_exact_week(block_issue_log, selected_week),
        weekly_action=filter_exact_week(block_action_tracker, selected_week),
        weekly_fuel=filter_exact_week(block_fuel_ratio, selected_week),
        weekly_hauling=filter_exact_week(block_hauling_review, selected_week),
        weekly_findings=filter_exact_week(block_findings, selected_week),
        weekly_inventory=filter_exact_week(block_inventory, selected_week),
        weekly_unit=filter_exact_week(block_unit_performance, selected_week), # Data terfilter block
        weekly_unit_full=weekly_unit_full,                                   # Data seluruh block (FIX 1)
        latest_df=latest_df, 
        prev_df=prev_df, 
        analysis_result=analysis_result,
        excel_func=to_excel
    )

else:
    max_date = block_df["week_date"].max()
    period_mode = st.sidebar.selectbox("Select Period Mode", ["MTD", "YTD", "Custom Date Range"])

    if period_mode == "Custom Date Range":
        custom_start = st.sidebar.date_input("Start Date", value=block_df["week_date"].min().date())
        custom_end = st.sidebar.date_input("End Date", value=max_date.date())
    else:
        custom_start, custom_end = None, None

    start_date, end_date = get_period_dates(period_mode, max_date, custom_start, custom_end)
    
    # Routing ke render fungsi Range Page (DENGAN PERBAIKAN PATCH FINAL)
    render_range_page(
        block=block, 
        max_date=max_date, 
        block_df=block_df,
        range_issue=filter_range(block_issue_log, start_date, end_date),
        range_action=filter_range(block_action_tracker, start_date, end_date),
        range_fuel=filter_range(block_fuel_ratio, start_date, end_date),
        range_hauling=filter_range(block_hauling_review, start_date, end_date),
        range_findings=filter_range(block_findings, start_date, end_date),
        range_unit=filter_range(block_unit_performance, start_date, end_date), # <--- REVISI TERBARU
        start_date=start_date, 
        end_date=end_date,
        excel_func=to_excel
    )

st.caption("Professional dashboard generated from integrated weekly mining data.")