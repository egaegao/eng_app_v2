import streamlit as st
import pandas as pd
import plotly.express as px
from dashboard_ui import section_title, style_status_table
from dashboard_helpers import (
    safe_metric, score_label, format_metric_name, 
    format_status_ui, get_period_dates
)
from dashboard_summary import (
    build_site_score, summarize_findings, summarize_actions, 
    aggregate_range_kpi
)
from dashboard_charts import plot_simple_bar
from dashboard_sections import render_trend_section, render_snapshot_section

def render_range_page(
    block, max_date, block_df,
    range_issue, range_action, range_fuel, range_hauling, range_findings,
    range_unit,   # Parameter Unit Performance
    start_date, end_date, excel_func
):
    # Header Utama
    st.title("📅 Mining Performance Range Analysis")
    st.caption(f"Mining Block: {block} | Period: {start_date.strftime('%d %b %Y')} to {end_date.strftime('%d %b %Y')}")

    # Memfilter data berdasarkan rentang waktu yang dipilih
    range_df = block_df[(block_df["week_date"] >= start_date) & (block_df["week_date"] <= end_date)].copy()
    range_agg = aggregate_range_kpi(range_df)

    if range_agg.empty:
        st.warning("Tidak ada data dalam rentang tanggal yang dipilih.")
        return

    # ==========================================
    # 1. RANGE EXECUTIVE OVERVIEW (FIXED VISUAL)
    # ==========================================
    section_title("🧭 Range Executive Overview")
    
    # Menggunakan st.container(border=True) untuk kartu yang rapi dan profesional
    with st.container(border=True):
        ob = safe_metric(range_agg, "overburden")
        coal = safe_metric(range_agg, "coal_getting")
        rain = safe_metric(range_agg, "rain_hours")
        
        lines = []
        if ob is not None:
            status_ob = "on-track" if ob["achievement_pct"] >= 95 else ("mendekati target" if ob["achievement_pct"] >= 80 else "masih rendah")
            lines.append(f"Overburden kumulatif {status_ob} di {ob['achievement_pct']:.1f}% dari plan.")
        
        if coal is not None:
            status_coal = "melampaui target" if coal["achievement_pct"] >= 100 else "tertinggal"
            lines.append(f"Coal getting kumulatif {status_coal} di {coal['achievement_pct']:.1f}% dari target.")
        
        if rain is not None and rain["achievement_pct"] > 100:
            lines.append(f"Rain hours kumulatif berada di atas plan ({rain['actual']:.1f} vs {rain['plan']:.1f} jam).")
            
        if not range_fuel.empty and "fr_achievement_pct" in range_fuel.columns:
            lines.append(f"Average fuel ratio achievement periode ini: {range_fuel['fr_achievement_pct'].mean():.1f}%.")

        if lines:
            summary_markdown = "\n".join([f"- {line}" for line in lines])
            st.markdown(summary_markdown)
        else:
            st.info("Pilih rentang waktu yang memiliki data untuk melihat ringkasan.")

    # Range Site Health Quick Stats
    range_score = build_site_score(range_agg)
    findings_count = summarize_findings(range_findings)
    action_counts = summarize_actions(range_action)

    section_title("🏥 Range Site Health")
    score_col1, score_col2, score_col3, score_col4 = st.columns(4)
    score_col1.metric("Range Health Score", f"{range_score:.1f}", score_label(range_score))
    score_col2.metric("High Issues", int((range_issue["severity"] == "High").sum()) if "severity" in range_issue.columns else 0)
    score_col3.metric("Open Actions", action_counts["Open"])
    score_col4.metric("Open Findings", findings_count["Open"])

    st.divider()

    # ==========================================
    # 2. RANGE KEY METRICS (BIG NUMBERS)
    # ==========================================
    section_title("🚀 Range Key Metrics")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    ewh = safe_metric(range_agg, "ewh_ob")
    ewh_coal = safe_metric(range_agg, "ewh_coal")
    sr = safe_metric(range_agg, "stripping_ratio")
    
    if ob is not None: 
        col1.metric("Overburden", f"{ob['actual']:,.0f} BCM", f"{ob['achievement_pct']:.1f}%")
    if coal is not None: 
        col2.metric("Coal Getting", f"{coal['actual']:,.0f} Ton", f"{coal['achievement_pct']:.1f}%")
    if rain is not None: 
        col3.metric("Rain Hours", f"{rain['actual']:.1f} Jam", f"{rain['actual'] - rain['plan']:+.1f}", delta_color="inverse")
    if ewh is not None: 
        col4.metric("EWH OB", f"{ewh['actual']:.1f} Jam", f"{ewh['achievement_pct']:.1f}%")
    if ewh_coal is not None:
        col5.metric("EWH Coal", f"{ewh_coal['actual']:.1f} Jam", f"{ewh_coal['achievement_pct']:.1f}%")
    if sr is not None:
        col6.metric("Stripping Ratio", f"{sr['actual']:.2f}", f"{sr['achievement_pct']:.1f}%")

    st.divider()

    # ==========================================
    # 3. KPI SUMMARY (STYLING FIX: st.table & COLUMN MATCH)
    # ==========================================
    section_title("📈 Range KPI Summary")
    
    # 1. Pastikan kolom 'metric' diformat
    # 2. Rename kolom 'status' menjadi 'Status' agar match dengan helper style_status_table
    show_df = range_agg[["metric", "plan", "actual", "achievement_pct", "gap", "status"]].copy()
    show_df["metric"] = show_df["metric"].apply(format_metric_name)
    show_df["status"] = show_df["status"].apply(format_status_ui)
    
    # Sinkronisasi Nama Kolom ke Capitalized 'Status'
    show_df = show_df.rename(columns={"status": "Status"})
    
    # Gunakan st.table() karena st.dataframe() sering mengabaikan style background CSS
    st.table(style_status_table(show_df))

    # ==========================================
    # 4. RANGE PERFORMANCE SNAPSHOT
    # ==========================================
    st.divider()
    section_title("📊 Range Performance Snapshot")
    render_snapshot_section(
        range_agg,
        ["overburden", "coal_getting", "rain_hours", "ewh_ob", "ewh_coal", "stripping_ratio"]
    )

    st.divider()

    # ==========================================
    # 5. RANGE FUEL ANALYSIS
    # ==========================================
    section_title("⛽ Range Fuel Analysis")
    if not range_fuel.empty:
        fuel_summary = range_fuel.groupby("activity", as_index=False)[["plan_liter", "actual_liter"]].sum() if set(["activity", "plan_liter", "actual_liter"]).issubset(range_fuel.columns) else pd.DataFrame()
        if not fuel_summary.empty:
            st.dataframe(fuel_summary, use_container_width=True, hide_index=True)
            fig_fuel = plot_simple_bar(fuel_summary, "activity", "plan_liter", "actual_liter", "Range Fuel Consumption", "Liter")
            st.plotly_chart(fig_fuel, use_container_width=True)
    else:
        st.info("Tidak ada data fuel ratio pada rentang ini.")

    st.divider()

    # ==========================================
    # 6. RANGE TREND
    # ==========================================
    section_title("📈 Range Trend")
    render_trend_section(
        range_df, 
        ["overburden", "coal_getting", "rain_hours", "ewh_ob", "ewh_coal", "stripping_ratio"]
    )

    st.divider()

    # ==========================================
    # 7. 🚜 RANGE UNIT PERFORMANCE (ADVANCED)
    # ==========================================
    section_title("🚜 Range Unit Performance")

    if range_unit is not None and not range_unit.empty:
        colf1, colf2 = st.columns(2)
        block_options = ["All", "Zebra", "Utara"]
        selected_block = colf1.selectbox("Filter Block (Unit - Range)", block_options, index=0)

        up = range_unit.copy()
        if selected_block != "All":
            up = up[up["block"].str.upper() == selected_block.upper()]

        cat_options = ["All"] + sorted(up["category"].dropna().unique().tolist())
        selected_cat = colf2.selectbox("Filter Category (Range)", cat_options)

        if selected_cat != "All":
            up = up[up["category"] == selected_cat]

        st.write("")

        if up.empty:
            st.warning("Tidak ada data unit pada range ini.")
            st.divider()
        else:
            # KPI Cards
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Unit", up["unit_id"].nunique())
            col2.metric("Avg PA", f"{up['pa_actual'].mean():.1f}%")
            col3.metric("Avg MA", f"{up['ma_actual'].mean():.1f}%")
            col4.metric("Avg UA", f"{up['ua_actual'].mean():.1f}%")

            st.write("")

            # Breakdown Logic
            if selected_cat == "All":
                group_col = "category"
                title_text = "Performance by Category"
            else:
                group_col = "type" if "type" in up.columns else "unit_type"
                title_text = f"Performance by Type (Category: {selected_cat})"

            st.markdown(f"**{title_text}**")

            cat_summary = (
                up.groupby(group_col)
                .agg(
                    total_unit=("unit_id", "nunique"),
                    avg_pa=("pa_actual", "mean"),
                    avg_ma=("ma_actual", "mean"),
                    avg_ua=("ua_actual", "mean")
                )
                .reset_index()
            )

            # Performance Table (Using Styler for better readability)
            st.dataframe(
                cat_summary.style
                .format({
                    "avg_pa": "{:.1f}",
                    "avg_ma": "{:.1f}",
                    "avg_ua": "{:.1f}"
                })
                .set_properties(**{
                    'font-size': '15px',
                    'color': 'black',
                    'font-weight': '600'
                }),
                use_container_width=True,
                hide_index=True
            )

            with st.expander("Detail Unit Performance", expanded=False):
                st.dataframe(
                    up.style.format({
                        "pa_plan": "{:,.1f}", "pa_actual": "{:,.1f}",
                        "ma_plan": "{:,.1f}", "ma_actual": "{:,.1f}",
                        "ua_plan": "{:,.1f}", "ua_actual": "{:,.1f}"
                    }, decimal=',', thousands='.')
                    .set_properties(**{
                        'font-size': '14px',
                        'color': 'black',
                        'font-weight': '600'
                    }),
                    use_container_width=True,
                    hide_index=True
                )

            st.write("")

            # Performance Chart
            chart_df = cat_summary.copy()
            chart_df = chart_df.melt(
                id_vars=group_col,
                value_vars=["avg_pa", "avg_ma", "avg_ua"],
                var_name="metric",
                value_name="value"
            )
            chart_df["metric"] = chart_df["metric"].replace({
                "avg_pa": "PA", "avg_ma": "MA", "avg_ua": "UA"
            })

            fig_unit = px.bar(
                chart_df,
                x=group_col,
                y="value",
                color="metric",
                barmode="group",
                text=chart_df["value"].round(1),
                title=title_text,
                color_discrete_sequence=["#FFA559", "#66BB6A", "#42A5F5"]
            )
            fig_unit.update_traces(textposition='outside')
            fig_unit.update_layout(
                font=dict(size=16, color="black"),
                xaxis=dict(tickfont=dict(size=14, color="black")),
                yaxis=dict(tickfont=dict(size=14, color="black")),
                legend=dict(font=dict(size=14)),
            )
            st.plotly_chart(fig_unit, use_container_width=True)

            # Top & Worst Ranking
            st.write("")
            st.markdown("**Top & Worst Unit Performance**")
            up["score"] = (up["pa_actual"] + up["ma_actual"] + up["ua_actual"]) / 3
            best_unit = up.sort_values("score", ascending=False).head(5)
            worst_unit = up.sort_values("score", ascending=True).head(5)

            colA, colB = st.columns(2)
            with colA:
                st.markdown("🏆 Top 5 Unit")
                st.dataframe(best_unit[["unit_id", "unit_type", "category"]], use_container_width=True, hide_index=True)
            with colB:
                st.markdown("⚠️ Worst 5 Unit")
                st.dataframe(worst_unit[["unit_id", "unit_type", "category"]], use_container_width=True, hide_index=True)

    else:
        st.info("Tidak ada data unit performance pada periode ini.")

    st.divider()

    # ==========================================
    # 8. DOWNLOAD SECTION
    # ==========================================
    if not range_agg.empty:
        excel_data = excel_func(range_agg)
        st.download_button(
            label="Download Range KPI Summary (Excel)", 
            data=excel_data, 
            file_name=f"range_kpi_{block}_{start_date.strftime('%Y%m%d')}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )