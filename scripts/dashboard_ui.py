import streamlit as st

def section_title(text):
    st.markdown("<div class='dashboard-section'></div>", unsafe_allow_html=True)
    st.subheader(text)

def style_status_table(df, status_column="status"):
    def row_style(row):
        # PATCH 2: Logika fleksibel untuk mengambil data dari kolom status
        # Mencoba mengambil 'status' (lowercase) atau 'Status' (capitalized)
        status = str(
            row.get(status_column) 
            or row.get(status_column.capitalize()) 
            or ""
        ).lower()
        
        # BEST PRACTICE UX: Background pastel dengan warna teks senada yang gelap dan tebal
        if any(keyword in status for keyword in ["red", "critical", "open"]):
            return ["background-color: #fee2e2; color: #991b1b; font-weight: 500;"] * len(row)
        elif any(keyword in status for keyword in ["yellow", "warning", "progress"]):
            return ["background-color: #ffedd5; color: #9a3412; font-weight: 500;"] * len(row)
        elif any(keyword in status for keyword in ["green", "good", "closed", "close"]):
            return ["background-color: #d1fae5; color: #064e3b; font-weight: 500;"] * len(row)
        return [""] * len(row)
    
    # BEST PRACTICE READABILITY: Format angka dengan koma ribuan dan 2 desimal
    return df.style.apply(row_style, axis=1).format(precision=2, thousands=",")

def style_issue_table(df):
    def row_style(row):
        category = str(row.get("Category", "")).strip().lower()
        if category == "weather":
            return ["background-color: #ffedd5; color: #9a3412; font-weight: 500;"] * len(row)
        elif category == "pumping":
            return ["background-color: #e0f2fe; color: #075985; font-weight: 500;"] * len(row)
        elif category == "fleet":
            return ["background-color: #fee2e2; color: #991b1b; font-weight: 500;"] * len(row)
        return [""] * len(row)
    
    # BEST PRACTICE READABILITY: Format angka dengan koma ribuan dan 2 desimal
    return df.style.apply(row_style, axis=1).format(precision=2, thousands=",")