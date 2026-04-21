import pandas as pd
from dashboard_config import LOWER_IS_BETTER

def safe_divide(a, b):
    if pd.isna(b) or b == 0:
        return 0.0
    return a / b

def parse_week_date(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "week_date" in df.columns:
        df["week_date"] = pd.to_datetime(
            df["week_date"],
            format="mixed",
            dayfirst=True,
            errors="coerce"
        )
    return df

def normalize_text(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace("_", " ", regex=False)
        .str.title()
    )

def get_status(metric, achievement_pct):
    if metric in LOWER_IS_BETTER:
        if achievement_pct <= 100:
            return "🟢 GOOD"
        elif achievement_pct <= 120:
            return "🟡 WARNING"
        return "🔴 CRITICAL"
    else:
        if achievement_pct >= 95:
            return "🟢 GOOD"
        elif achievement_pct >= 80:
            return "🟡 WARNING"
        return "🔴 CRITICAL"

def score_label(score):
    if score >= 95:
        return "🟢 Strong"
    elif score >= 80:
        return "🟡 Watch"
    return "🔴 Critical"

def format_metric_name(metric):
    mapping = {
        "overburden": "Overburden",
        "coal_getting": "Coal Getting",
        "rain_hours": "Rain Hours",
        "slippery_hours": "Slippery Hours",
        "ewh_ob": "EWH OB",
        "ewh_coal": "EWH Coal",
        "fuel_ratio_mining": "Fuel Ratio",
        "stripping_ratio": "Stripping Ratio",
    }
    return mapping.get(metric, metric)

def format_status_ui(status):
    status_str = str(status).upper().strip()
    if status_str == "GREEN":
        return "🟢 GOOD"
    elif status_str == "YELLOW":
        return "🟡 WARNING"
    elif status_str == "RED":
        return "🔴 CRITICAL"
    return status

def safe_metric(df, metric_name):
    if df.empty or "metric" not in df.columns:
        return None
    sub = df[df["metric"] == metric_name]
    if sub.empty:
        return None
    return sub.iloc[0]

def limit_rows(df, option):
    if option == "All":
        return df.copy()
    return df.head(int(option))

def get_period_dates(period_mode, max_date, custom_start=None, custom_end=None):
    max_ts = pd.Timestamp(max_date)

    if period_mode == "MTD":
        start_date = max_ts.replace(day=1)
        end_date = max_ts
    elif period_mode == "YTD":
        start_date = max_ts.replace(month=1, day=1)
        end_date = max_ts
    else:
        start_date = pd.Timestamp(custom_start)
        end_date = pd.Timestamp(custom_end)

    return start_date, end_date