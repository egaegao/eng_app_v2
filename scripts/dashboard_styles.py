import streamlit as st

def apply_custom_styles():
    st.markdown("""
    <style>
    /* ========================================= */
    /* 1. GLOBAL CONTAINER                       */
    /* ========================================= */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1550px;
    }

    /* ========================================= */
    /* 2. MINIMALIST METRIC CARDS (3 ROWS)       */
    /* ========================================= */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        /* Aksen garis kiri tipis & profesional */
        border-left: 4px solid #475569 !important; 
        border: 1px solid #e2e8f0;
        
        padding: 15px 20px !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        
        display: flex;
        flex-direction: column !important;
        align-items: flex-start !important;
    }

    /* Memaksa susunan vertikal */
    div[data-testid="stMetric"] > div {
        display: flex !important;
        flex-direction: column !important;
        align-items: flex-start !important;
        width: 100% !important;
        gap: 2px !important;
    }

    /* Baris 1: Judul Metrik */
    div[data-testid="stMetricLabel"] p {
        font-size: 0.95rem !important;
        font-weight: 500 !important;
        color: #64748b !important;
        margin: 0 !important;
    }

    /* Baris 2: Angka Utama */
    div[data-testid="stMetricValue"] > div {
        font-size: 1.8rem !important; /* Ukuran font dirampingkan */
        font-weight: 700 !important;
        color: #1e293b !important;
        line-height: 1.2 !important;
    }

    /* Baris 3: Status Delta */
    div[data-testid="stMetricDelta"] {
        margin-top: 2px !important;
    }

    div[data-testid="stMetricDelta"] > div {
        font-size: 0.9rem !important;
        font-weight: 600 !important;
        display: inline-flex !important;
        padding: 2px 8px !important;
        border-radius: 4px !important;
        background-color: rgba(0,0,0,0.02) !important;
    }

    /* Ukuran Icon Panah */
    div[data-testid="stMetricDelta"] svg {
        width: 14px !important;
        height: 14px !important;
    }

    /* ========================================= */
    /* 3. CLEAN EXECUTIVE CONTAINER              */
    /* ========================================= */
    div[data-testid="stElementContainer"] div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #ffffff;
        border-radius: 8px !important;
        border: 1px solid #e2e8f0 !important;
        border-left: 4px solid #1e293b !important; /* Aksen gelap untuk pembeda */
    }

    /* ========================================= */
    /* 4. UTILITY STYLES                         */
    /* ========================================= */
    .detail-title {
        font-weight: 600;
        font-size: 1rem;
        color: #334155;
        margin-bottom: 8px;
    }
    
    hr {
        margin-top: 1rem !important;
        margin-bottom: 1rem !important;
        border-top: 1px solid #f1f5f9 !important;
    }

    /* Sidebar Clean Look */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    </style>
    """, unsafe_allow_html=True)