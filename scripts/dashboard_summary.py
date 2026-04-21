import pandas as pd
from dashboard_config import LOWER_IS_BETTER, PRIMARY_METRICS
from dashboard_helpers import safe_divide, safe_metric, get_status, score_label

def build_site_score(df):
    """Menghitung skor kesehatan site berdasarkan pencapaian KPI."""
    if df.empty:
        return 0.0

    score_df = df.copy()
    if "achievement_pct" not in score_df.columns:
        return 0.0

    # Logika skor: metrik 'Lower Is Better' (seperti Rain) dibalik perhitungannya
    score_df["score_component"] = score_df.apply(
        lambda r: max(0, min(120, 200 - r["achievement_pct"]))
        if r["metric"] in LOWER_IS_BETTER
        else max(0, min(120, r["achievement_pct"])),
        axis=1
    )

    focus = score_df[score_df["metric"].isin(PRIMARY_METRICS)].copy()
    if focus.empty:
        focus = score_df.copy()

    return round(focus["score_component"].mean(), 1)

def summarize_actions(actions_df):
    """Menghitung ringkasan status pada Action Tracker."""
    if actions_df.empty or "status" not in actions_df.columns:
        return {"Open": 0, "Progress": 0, "Closed": 0}

    counts = actions_df["status"].astype(str).str.strip().str.title().value_counts().to_dict()
    return {
        "Open": counts.get("Open", 0),
        "Progress": counts.get("Progress", 0),
        "Closed": counts.get("Close", 0) + counts.get("Closed", 0),
    }

def summarize_findings(findings_df):
    """Menghitung ringkasan temuan (Findings) dari data detail."""
    
    if findings_df.empty or "status" not in findings_df.columns:
        return {"Open": 0, "Progress": 0, "Closed": 0, "Total": 0}

    # Menggunakan value_counts() untuk menghitung jumlah baris per status
    status_series = findings_df["status"].astype(str).str.strip().str.title()
    counts = status_series.value_counts()

    return {
        "Open": int(counts.get("Open", 0)),
        "Progress": int(counts.get("Progress", 0)),
        "Closed": int(counts.get("Closed", 0)),
        "Total": int(len(findings_df))
    }

def aggregate_range_kpi(df):
    """Aggregasi data KPI untuk periode Range Analysis."""
    if df.empty:
        return pd.DataFrame()

    grouped = df.groupby(["block", "metric"], as_index=False)[["plan", "actual"]].sum()
    grouped["achievement_pct"] = grouped.apply(lambda r: round(safe_divide(r["actual"], r["plan"]) * 100, 1), axis=1)
    grouped["status"] = grouped.apply(lambda r: get_status(r["metric"], r["achievement_pct"]), axis=1)
    grouped["gap"] = (grouped["actual"] - grouped["plan"]).round(2)
    
    def _priority(row):
        if row["metric"] in LOWER_IS_BETTER:
            return max(0, row["achievement_pct"] - 100)
        return max(0, 100 - row["achievement_pct"])
        
    grouped["priority_score"] = grouped.apply(_priority, axis=1)
    return grouped

def build_executive_summary(latest_df, prev_df, issue_df, action_df, fuel_df, haul_df, findings_df):
    """Membangun narasi analisis mingguan yang dinamis."""
    lines = []
    
    # 1. Analisis Skor Kesehatan Site
    site_score = build_site_score(latest_df)
    lines.append(f"Kesehatan operasional site berada pada level **{score_label(site_score)}** ({site_score:.1f} poin).")

    # 2. Analisis KPI Produksi
    ob = safe_metric(latest_df, "overburden")
    coal = safe_metric(latest_df, "coal_getting")
    rain = safe_metric(latest_df, "rain_hours")
    ewh = safe_metric(latest_df, "ewh_ob")

    if ob is not None:
        if ob["achievement_pct"] < 80:
            lines.append(f"Produksi Overburden masih kritis di level {ob['achievement_pct']:.1f}% dari plan.")
        elif ob["achievement_pct"] < 95:
            lines.append(f"Overburden mendekati target dengan pencapaian {ob['achievement_pct']:.1f}%.")
        else:
            lines.append(f"Overburden berkinerja baik (On-Track) di level {ob['achievement_pct']:.1f}%.")

    if coal is not None:
        if coal["achievement_pct"] >= 100:
            lines.append(f"Coal getting melampaui target mingguan ({coal['achievement_pct']:.1f}%).")
        elif coal["achievement_pct"] < 80:
            lines.append(f"Coal getting tertinggal signifikan pada {coal['achievement_pct']:.1f}% dari target.")

    # 3. Analisis Faktor Eksternal & Kendala
    if rain is not None and rain["actual"] > rain["plan"]:
        lines.append(f"Hujan di atas rencana ({rain['actual']:.1f} vs {rain['plan']:.1f} jam) menjadi pembatas utama produktivitas.")

    if ewh is not None and ewh["achievement_pct"] < 80:
        lines.append(f"EWH OB rendah ({ewh['achievement_pct']:.1f}%), mengindikasikan utilisasi alat muat belum optimal.")

    # 4. Analisis Perbandingan Week-on-Week (WoW)
    if prev_df is not None and not prev_df.empty:
        prev_ob = safe_metric(prev_df, "overburden")
        prev_coal = safe_metric(prev_df, "coal_getting")

        if ob is not None and prev_ob is not None:
            diff = ob["actual"] - prev_ob["actual"]
            word = "naik" if diff > 0 else "turun"
            lines.append(f"Volume OB {word} {abs(diff):,.0f} BCM dibandingkan minggu lalu.")

        if coal is not None and prev_coal is not None:
            if coal["actual"] < prev_coal["actual"]:
                lines.append(f"Terjadi pelemahan coal getting sebesar {abs(coal['actual'] - prev_coal['actual']):,.0f} ton WoW.")

    # 5. Ringkasan Isu & Tindak Lanjut
    if not issue_df.empty and "issue_category" in issue_df.columns:
        top_issue = issue_df["issue_category"].value_counts().idxmax()
        lines.append(f"Isu operasional dominan: **{top_issue}**.")

    action_count = summarize_actions(action_df)
    if action_count["Open"] > 0:
        lines.append(f"Status Action Tracker: {action_count['Open']} Open, {action_count['Progress']} In Progress.")

    return lines

def build_recommendations(latest_df, prev_df, issue_df, action_df, fuel_df, haul_df, findings_df):
    """Menghasilkan saran tindakan berdasarkan deviasi data."""
    recs = []
    ob = safe_metric(latest_df, "overburden")
    rain = safe_metric(latest_df, "rain_hours")
    ewh = safe_metric(latest_df, "ewh_ob")

    if rain is not None and rain["actual"] > rain["plan"]:
        recs.append("Maksimalkan weather window dan siapkan recovery plan drainage.")
    
    if ewh is not None and ewh["achievement_pct"] < 85:
        recs.append("Audit ketersediaan operator dan breakdown unit untuk memulihkan EWH.")

    if ob is not None and ob["achievement_pct"] < 90:
        recs.append("Evaluasi stripping ratio area kerja dan prioritaskan expose coal.")

    if not action_df.empty and "overdue_flag" in action_df.columns:
        if action_df["overdue_flag"].any():
            recs.append("Percepat penyelesaian item action tracker yang sudah Overdue.")

    if not recs:
        recs.append("Pertahankan ritme operasional dan monitor deviasi KPI harian secara ketat.")

    return recs