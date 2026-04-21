import streamlit as st
import pandas as pd
import plotly.express as px
# Import module internal (Asumsi file-file ini ada di direktori project D:\opsapp)
from dashboard_ui import section_title, style_status_table, style_issue_table
from dashboard_helpers import (
    safe_metric, score_label, format_metric_name, 
    format_status_ui, limit_rows
)
from dashboard_summary import (
    build_executive_summary, build_site_score, summarize_findings, 
    summarize_actions, build_recommendations
)
from dashboard_charts import plot_simple_bar, plot_inventory_bar
from dashboard_sections import render_snapshot_section, render_trend_section

def format_idn(val, precision=2):
    """Helper internal untuk format angka Indonesia di kartu metrik"""
    if val is None or pd.isna(val): return "0"
    # Format ribuan dengan koma, lalu tukar koma jadi titik dan titik jadi koma
    formatted = f"{val:,.{precision}f}"
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")

def render_weekly_page(
    block, selected_week, available_weeks, block_df,
    weekly_issue, weekly_action, weekly_fuel, weekly_hauling, weekly_findings,
    weekly_inventory, weekly_unit, weekly_unit_full, 
    latest_df, prev_df, analysis_result, excel_func
):
    # Header Utama
    st.title("📊 Weekly Mining Dashboard")
    
    # Header Sub-informasi (Block & Tanggal)
    st.markdown(f"""
        <p style="font-size: 1.1rem; color: #334155; font-weight: 500; margin-top: -15px; margin-bottom: 20px;">
            Mining Block: <b style="color: #1e293b;">{block}</b> | Week: <b style="color: #1e293b;">{pd.Timestamp(selected_week).strftime('%d %b %Y')}</b>
        </p>
    """, unsafe_allow_html=True)

    # ==========================================
    # 1. EXECUTIVE OVERVIEW
    # ==========================================
    section_title("🧭 Executive Overview")
    
    with st.container(border=True):
        summary_lines = build_executive_summary(
            latest_df, prev_df, weekly_issue, weekly_action, 
            weekly_fuel, weekly_hauling, weekly_findings
        )
        
        if summary_lines:
            summary_markdown = "\n".join([f"- {line}" for line in summary_lines])
            st.markdown(summary_markdown)
        else:
            st.info("Pilih periode minggu yang valid untuk melihat analisis operasional.")

    # Quick Stats Metrics
    site_score = build_site_score(latest_df)
    score_text = score_label(site_score)
    findings_count = summarize_findings(weekly_findings)
    action_counts = summarize_actions(weekly_action)

    score_col1, score_col2, score_col3, score_col4 = st.columns(4)
    # Metrik dengan format desimal koma (Lokalisasi IDN)
    score_col1.metric("Site Health Score", format_idn(site_score, 1), score_text)
    score_col2.metric("High Issues", int((weekly_issue["severity"] == "High").sum()) if "severity" in weekly_issue.columns else 0)
    score_col3.metric("Open Actions", action_counts["Open"])
    score_col4.metric("Open Findings", findings_count["Open"])

    st.divider()

    # ==========================================
    # 2. KEY METRICS (BIG NUMBERS)
    # ==========================================
    section_title("🚀 Key Metrics")
    
    # Pengambilan metrik termasuk Stripping Ratio
    ob = safe_metric(latest_df, "overburden")
    coal = safe_metric(latest_df, "coal_getting")
    rain = safe_metric(latest_df, "rain_hours")
    ewh = safe_metric(latest_df, "ewh_ob")
    ewh_coal = safe_metric(latest_df, "ewh_coal")
    sr = safe_metric(latest_df, "stripping_ratio")

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    if ob is not None:
        col1.metric("Overburden", f"{format_idn(ob['actual'], 0)} BCM", f"{format_idn(ob['achievement_pct'], 1)}% vs plan")
    if coal is not None:
        col2.metric("Coal Getting", f"{format_idn(coal['actual'], 0)} Ton", f"{format_idn(coal['achievement_pct'], 1)}% vs plan")
    if rain is not None:
        col3.metric("Rain Hours", f"{format_idn(rain['actual'], 1)} Jam", f"{format_idn(rain['actual'] - rain['plan'], 1)} vs plan", delta_color="inverse")
    if ewh is not None:
        col4.metric("EWH OB", f"{format_idn(ewh['actual'], 1)} Jam", f"{format_idn(ewh['achievement_pct'], 1)}% vs plan")

    if ewh_coal is not None:
        col5.metric(
            "EWH Coal",
            f"{format_idn(ewh_coal['actual'], 1)} Jam",
            f"{format_idn(ewh_coal['achievement_pct'], 1)}% vs plan"
        )

    if sr is not None:
        col6.metric(
            "Stripping Ratio (SR)",
            f"{format_idn(sr['actual'], 2)}",
            f"{format_idn(sr['achievement_pct'], 1)}% vs plan"
        )

    st.divider()

    # ==========================================
    # 3. KPI TABLES (FULL WIDTH)
    # ==========================================
    # Bagian ini sudah diubah dari columns menjadi single container
    kpi_container = st.container()

    section_title("📈 KPI Summary")
    show_df = latest_df[["week_date", "metric", "plan", "actual", "achievement_pct", "gap", "status"]].copy()
    
    if "week_date" in show_df.columns:
        show_df["week_date"] = pd.to_datetime(show_df["week_date"]).dt.date

    show_df["metric"] = show_df["metric"].apply(format_metric_name)
    show_df["status"] = show_df["status"].apply(format_status_ui)
    
    # Terapkan Lokalisasi Indonesia pada Dataframe dengan lebar penuh
    st.dataframe(
        style_status_table(show_df, "status").format(
            {"plan": "{:,.2f}", "actual": "{:,.2f}", "achievement_pct": "{:,.2f}", "gap": "{:,.2f}"},
            decimal=',', thousands='.'
        ), 
        use_container_width=True, hide_index=True
    )

    st.divider()

    # ==========================================
    # 4. ISSUE ANALYSIS
    # ==========================================
    section_title("🚩 Issue Analysis")
    
    issue_df = weekly_issue.copy() if not weekly_issue.empty else pd.DataFrame()

    if not issue_df.empty:
        colf1, colf2 = st.columns(2)

        severity_filter = colf1.selectbox(
            "Filter Severity",
            ["All", "High", "Medium", "Low"],
            key="severity_filter"
        )

        category_filter = colf2.selectbox(
            "Filter Category",
            ["All"] + sorted(issue_df["issue_category"].dropna().unique().tolist()),
            key="category_filter"
        )

        # Apply Filter
        if severity_filter != "All":
            issue_df = issue_df[
                issue_df["severity"].astype(str).str.upper() == severity_filter.upper()
            ]

        if category_filter != "All":
            issue_df = issue_df[
                issue_df["issue_category"] == category_filter
            ]

    if not issue_df.empty:
        issue_summary = issue_df["issue_category"].value_counts().reset_index()
        issue_summary.columns = ["Issue Category", "Count"]
        st.dataframe(issue_summary.style.format(thousands="."), use_container_width=True, hide_index=True)

        with st.container(border=True):
            st.markdown("**Detailed Issues**")
            issue_view_option = st.selectbox("Show rows (Issues)", ["5", "10", "20", "All"], index=0, key="issue_view_option")
            detail_issues = issue_df[["week_date", "issue_category", "issue_detail", "impact_area", "status", "pic"]].copy()
            if "week_date" in detail_issues.columns:
                detail_issues["week_date"] = pd.to_datetime(detail_issues["week_date"]).dt.date
            detail_issues.columns = ["Date", "Category", "Issue Detail", "Impact Area", "Status", "PIC"]
            detail_issues = limit_rows(detail_issues, issue_view_option)
            st.dataframe(style_issue_table(detail_issues), use_container_width=True, hide_index=True)
    else:
        st.info("Tidak ada issue untuk filter yang dipilih.")

    st.write("") 

    section_title("🛠️ Action Tracker")
    actions = weekly_action.copy()
    if not actions.empty:
        action_summary = actions["status"].astype(str).str.title().value_counts().reset_index()
        action_summary.columns = ["Status", "Count"]
        st.dataframe(action_summary.style.format(thousands="."), use_container_width=True, hide_index=True)

        with st.container(border=True):
            st.markdown("**Detailed Action Tracker**")
            action_view_option = st.selectbox("Show rows (Actions)", ["5", "10", "20", "All"], index=0, key="action_view_option")
            actions_display = actions.copy()
            if "week_date" in actions_display.columns:
                actions_display["week_date"] = pd.to_datetime(actions_display["week_date"]).dt.date
            if "due_date" in actions_display.columns:
                actions_display["due_date"] = pd.to_datetime(actions_display["due_date"]).dt.date
            actions_display = limit_rows(actions_display, action_view_option)
            st.dataframe(
                style_status_table(actions_display, "status").format(decimal=',', thousands='.'), 
                use_container_width=True, hide_index=True
            )
    else:
        st.info("Tidak ada action tracker untuk periode ini.")

    st.divider()

    # ==========================================
    # 5. OPERATIONAL DETAILS (FUEL & INVENTORY)
    # ==========================================
    section_title("⛽ Fuel Analysis")
    if not weekly_fuel.empty:
        fuel_col1, fuel_col2, fuel_col3 = st.columns(3)
        fuel_col1.metric("Average Actual FR", format_idn(weekly_fuel['actual_fr'].mean()))
        fuel_col2.metric("Average Plan FR", format_idn(weekly_fuel['plan_fr'].mean()))
        fuel_col3.metric("Liter Deviation", format_idn(weekly_fuel['liter_dev'].sum(), 0))
        
        fuel_table = weekly_fuel.copy()
        if "week_date" in fuel_table.columns:
            fuel_table["week_date"] = pd.to_datetime(fuel_table["week_date"]).dt.date
        
        st.dataframe(
            style_status_table(fuel_table, "fuel_status").format(decimal=',', thousands='.'), 
            use_container_width=True, hide_index=True
        )
        st.plotly_chart(plot_simple_bar(weekly_fuel, "activity", "plan_liter", "actual_liter", "Fuel Plan vs Actual", "Liter"), use_container_width=True)
    else:
        st.info("Data fuel tidak tersedia untuk minggu ini.")

    st.divider()

    if not weekly_inventory.empty:
        section_title("📦 Inventory Monitoring")
        inv_col1, inv_col2 = st.columns(2)
        inv_col1.metric("Total Inventory", f"{format_idn(weekly_inventory['volume_ton'].sum(), 0)} Ton")
        inv_col2.metric("Inventory Types", weekly_inventory["inventory_type"].nunique())
        inv_summary = weekly_inventory.groupby(["block", "inventory_type"], as_index=False)["volume_ton"].sum()
        st.dataframe(inv_summary.style.format({"volume_ton": "{:,.2f}"}, decimal=',', thousands='.'), use_container_width=True, hide_index=True)
        st.plotly_chart(plot_inventory_bar(inv_summary), use_container_width=True)

    st.divider()

    section_title("🚛 Hauling Performance")
    if not weekly_hauling.empty:
        haul_col1, haul_col2, haul_col3 = st.columns(3)
        haul_col1.metric("Avg Achievement", f"{format_idn(weekly_hauling['achievement_pct'].mean(), 1)}%")
        haul_col2.metric("Target Ton", format_idn(weekly_hauling['target_ton'].sum(), 0))
        haul_col3.metric("Actual Ton", format_idn(weekly_hauling['actual_ton'].sum(), 0))
        
        haul_table = weekly_hauling.copy()
        if "week_date" in haul_table.columns:
            haul_table["week_date"] = pd.to_datetime(haul_table["week_date"]).dt.date
        st.dataframe(
            style_status_table(haul_table, "status").format(decimal=',', thousands='.'), 
            use_container_width=True, hide_index=True
        )
        st.plotly_chart(plot_simple_bar(weekly_hauling, "route", "target_ton", "actual_ton", "Hauling Performance", "Ton"), use_container_width=True)

    st.divider()

    # ==========================================
    # 🚜 UNIT PERFORMANCE
    # ==========================================
    section_title("🚜 Unit Performance")

    if weekly_unit is not None and not weekly_unit.empty:
        colf1, colf2 = st.columns(2)
        block_options = ["All", "Zebra", "Utara"]
        selected_block = colf1.selectbox("Filter Block (Unit)", block_options, index=0)
        
        up = weekly_unit.copy()
        if selected_block == "All":
            up = weekly_unit_full.copy()
        elif selected_block in ["Zebra", "Utara"]:
            up = weekly_unit_full[
                weekly_unit_full["block"].str.strip().str.upper() == selected_block.upper()
            ].copy()

        cat_options = ["All"] + sorted(up["category"].dropna().unique().tolist())
        selected_cat = colf2.selectbox("Filter Category", cat_options)

        if selected_cat != "All":
            up = up[up["category"] == selected_cat]

        st.write("")

        if not up.empty:
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Unit", up["unit_id"].nunique())
            col2.metric("Avg PA", f"{format_idn(up['pa_actual'].mean(), 1)}%")
            col3.metric("Avg MA", f"{format_idn(up['ma_actual'].mean(), 1)}%")
            col4.metric("Avg UA", f"{format_idn(up['ua_actual'].mean(), 1)}%")

            st.write("")

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
                labels={
                    group_col: group_col.upper(),
                    "value": "Percentage (%)",
                    "metric": "KPI"
                },
                color_discrete_sequence=["#FFA559", "#66BB6A", "#42A5F5"]
            )
            
            fig_unit.update_traces(textposition='outside')
            fig_unit.update_layout(
                font=dict(size=16, color="black"),
                xaxis=dict(tickfont=dict(size=14, color="black"), title_font=dict(size=16, color="black")),
                yaxis=dict(tickfont=dict(size=14, color="black"), title_font=dict(size=16, color="black")),
                legend=dict(font=dict(size=14)),
                margin=dict(t=50, b=50)
            )
            st.plotly_chart(fig_unit, use_container_width=True)

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
            st.warning("Tidak ada data unit untuk filter yang dipilih.")
    else:
        st.info("Data unit performance tidak tersedia.")

    st.divider()

    # ==========================================
    # 6. FINDINGS & TRENDS
    # ==========================================
    section_title("📋 Findings Summary")
    if not weekly_findings.empty:
        f_counts = summarize_findings(weekly_findings)
        fc1, fc2, fc3, fc4 = st.columns(4)
        fc1.metric("Open", f_counts["Open"])
        fc2.metric("Progress", f_counts["Progress"])
        fc3.metric("Closed", f_counts["Closed"])
        fc4.metric("Total", f_counts["Total"])
        
        with st.expander("Lihat Detail Findings", expanded=False):
            findings_display = weekly_findings.copy()
            if "week_date" in findings_display.columns:
                findings_display["week_date"] = pd.to_datetime(findings_display["week_date"]).dt.date
            if "due_date" in findings_display.columns:
                findings_display["due_date"] = pd.to_datetime(findings_display["due_date"]).dt.date

            if "category" in findings_display.columns and "status" in findings_display.columns:
                summary_cat = (
                    findings_display
                    .assign(status=findings_display["status"].astype(str).str.title())
                    .groupby(["category", "status"])
                    .size().unstack(fill_value=0).reset_index()
                )

                if not summary_cat.empty:
                    for s in ["Open", "Progress", "Closed"]:
                        if s not in summary_cat.columns: summary_cat[s] = 0
                    summary_cat["Total"] = summary_cat["Open"] + summary_cat["Progress"] + summary_cat["Closed"]
                    st.markdown("**Summary per Category**")
                    st.dataframe(summary_cat, use_container_width=True, hide_index=True)

            def highlight_status_row(row):
                status = str(row.get("status", "")).strip().title()
                if status == "Open": return ["background-color: #fee2e2"] * len(row)
                elif status == "Progress": return ["background-color: #fef9c3"] * len(row)
                elif status == "Closed": return ["background-color: #dcfce7"] * len(row)
                return [""] * len(row)

            st.dataframe(
                findings_display.style.apply(highlight_status_row, axis=1).format(thousands="."),
                use_container_width=True, hide_index=True
            )

    st.divider()

    section_title("📊 Performance Snapshot & Trend")
    # Snapshot Section
    render_snapshot_section(
        latest_df, 
        ["overburden", "coal_getting", "rain_hours", "ewh_ob", "ewh_coal", "stripping_ratio"]
    )
    
    # Trend Section
    render_trend_section(
        block_df, 
        ["overburden", "coal_getting", "rain_hours", "ewh_ob", "ewh_coal", "stripping_ratio"]
    )

    st.divider()

    # ==========================================
    # 7. RCA & RECOMMENDATIONS
    # ==========================================
    section_title("🧠 Root Cause & Recommendations")
    
    if prev_df is not None and not prev_df.empty:
        rain_prev = safe_metric(prev_df, "rain_hours")
        rain_now = safe_metric(latest_df, "rain_hours")
        ob_prev = safe_metric(prev_df, "overburden")
        ob_now = safe_metric(latest_df, "overburden")
        
        if (rain_now is not None) and (rain_prev is not None) and (ob_now is not None) and (ob_prev is not None):
            if rain_now["actual"] > rain_prev["actual"] and ob_now["actual"] < ob_prev["actual"]:
                st.error(f"⚠️ Penurunan OB ({format_idn(ob_now['actual'], 0)}) berkorelasi dengan kenaikan jam hujan ({format_idn(rain_now['actual'], 1)} jam).")

    recs = build_recommendations(latest_df, prev_df, weekly_issue, weekly_action, weekly_fuel, weekly_hauling, weekly_findings)
    for rec in recs:
        st.write(f"📌 {rec}")

    st.divider()

    # Download Button
    if not latest_df.empty:
        excel_data = excel_func(latest_df)
        st.download_button(
            label="Download Weekly Report (Excel)", 
            data=excel_data, 
            file_name=f"report_{block}_{pd.Timestamp(selected_week).strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )