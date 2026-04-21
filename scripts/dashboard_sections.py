import streamlit as st
import pandas as pd
from dashboard_charts import plot_snapshot_chart, plot_trend_chart
from dashboard_helpers import format_metric_name

def render_snapshot_section(data_df, metrics):
    for i in range(0, len(metrics), 2):
        row_cols = st.columns(2)
        for j, metric in enumerate(metrics[i:i+2]):
            sub = data_df[data_df["metric"] == metric].copy()
            if not sub.empty:
                sub_chart = sub[["metric", "plan", "actual"]].copy()
                sub_chart["metric"] = sub_chart["metric"].apply(format_metric_name)
                fig = plot_snapshot_chart(sub_chart, format_metric_name(metric))
                with row_cols[j]:
                    st.plotly_chart(fig, use_container_width=True)

def render_trend_section(data_df, metrics):
    # --- LOGIKA BATCHING (PER 26 MINGGU) ---
    all_weeks = sorted(data_df["week_date"].unique())
    total_weeks = len(all_weeks)
    limit = 26  
    
    display_df = data_df.copy()
    
    if total_weeks > limit:
        st.write("")
        n_batches = (total_weeks + limit - 1) // limit
        options = []
        batch_map = {}
        
        for n in range(n_batches):
            start_idx = n * limit
            end_idx = min((n + 1) * limit, total_weeks)
            
            date_start = all_weeks[start_idx]
            date_end = all_weeks[end_idx - 1]
            
            # KETERANGAN BERSIH: Hanya Tanggal Awal - Tanggal Akhir
            label = f"{pd.Timestamp(date_start).strftime('%d %b %Y')} – {pd.Timestamp(date_end).strftime('%d %b %Y')}"
            options.append(label)
            batch_map[label] = (date_start, date_end)
        
        # Default (index=len(options)-1) memastikan data TERBARU yang muncul pertama kali
        selected_range = st.selectbox(
            "Pilih Rentang Waktu Tren:", 
            options, 
            index=len(options)-1 
        )
        
        t_start, t_end = batch_map[selected_range]
        display_df = data_df[(data_df["week_date"] >= t_start) & (data_df["week_date"] <= t_end)].copy()

    # --- RENDER GRAFIK ---
    for i in range(0, len(metrics), 2):
        row_cols = st.columns(2)
        for j, metric in enumerate(metrics[i:i+2]):
            sub = display_df[display_df["metric"] == metric].sort_values("week_date").copy()
            if not sub.empty:
                fig = plot_trend_chart(sub, format_metric_name(metric))
                with row_cols[j]:
                    st.plotly_chart(fig, use_container_width=True)