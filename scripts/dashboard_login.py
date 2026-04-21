import streamlit as st

def login_screen():
    """Menampilkan layar login ultra-minimalis tanpa card redundant"""
    
    # CSS Kustom untuk membersihkan tampilan
    st.markdown("""
        <style>
        /* Menghilangkan elemen default Streamlit */
        header {visibility: hidden;}
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* Judul Utama */
        .main-title {
            color: #0f172a;
            font-size: 2.8rem;
            font-weight: 800;
            text-align: center;
            margin-top: 20px;
            margin-bottom: 25px;
            letter-spacing: -0.05em;
        }
        
        /* Kutipan Bijak */
        .quote-text {
            text-align: center;
            max-width: 600px;
            margin: 0 auto 10px auto;
            line-height: 1.6;
            color: #334155;
            font-size: 1.1rem;
            font-style: italic;
        }

        /* Nama Author */
        .author-text {
            color: #64748b;
            font-size: 0.8rem;
            text-align: center;
            margin-bottom: 45px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.2em;
        }

        /* Label Access Code */
        .input-label {
            font-size: 0.85rem;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 8px;
            text-align: left;
        }

        /* Tombol Modern & Tegas */
        div.stButton > button {
            background-color: #0f172a !important;
            color: #ffffff !important;
            border-radius: 6px !important;
            border: none !important;
            height: 45px !important;
            font-weight: 600 !important;
            width: 100% !important;
            margin-top: 10px !important;
        }
        
        /* Hilangkan padding container agar tidak muncul shadow/border sisa */
        div[data-testid="stVerticalBlock"] > div:has(div.login-area) {
            background-color: transparent !important;
            box-shadow: none !important;
            border: none !important;
        }
        </style>
    """, unsafe_allow_html=True)

    if 'authenticated' not in st.session_state:
        st.session_state['authenticated'] = False

    if not st.session_state['authenticated']:
        # Sembunyikan sidebar
        st.markdown("<style>section[data-testid='stSidebar'] {display: none;}</style>", unsafe_allow_html=True)
        
        # 1. Judul
        st.markdown('<h1 class="main-title">Mining System</h1>', unsafe_allow_html=True)
        
        # 2. Kutipan Bijak
        st.markdown("""
            <div class="quote-text">
                "Jika kamu kalah dalam kecerdasan, menanglah dalam ketekunan—itulah pertarungan yang menentukan. 
                Karena dalam setiap langkah sabar dan usaha, ada pertolongan Allah yang diam-diam menguatkan 
                dan meninggikan derajatmu."
            </div>
        """, unsafe_allow_html=True)
        
        # 3. Penulis
        st.markdown('<div class="author-text">Indrawijaya, CEE</div>', unsafe_allow_html=True)
        
        # 4. Input Section (Tanpa wrapper card)
        _, col_card, _ = st.columns([1.3, 1, 1.3])
        
        with col_card:
            # Menggunakan div class untuk identifikasi CSS jika perlu, tapi tanpa styling border/background
            st.markdown('<div class="login-area">', unsafe_allow_html=True)
            st.markdown('<div class="input-label">Access Code</div>', unsafe_allow_html=True)
            
            password = st.text_input("Access Code", type="password", placeholder="••••••••", label_visibility="collapsed")
            
            if st.button("Unlock Access"):
                if password == "pms":
                    st.session_state['authenticated'] = True
                    st.rerun()
                else:
                    st.error("Invalid Code.")
            st.markdown('</div>', unsafe_allow_html=True)
            
        st.stop()