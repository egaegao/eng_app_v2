import pandas as pd

def filter_by_block(df, block_value):
    if "block" not in df.columns:
        return df.copy()
    return df[df["block"].astype(str) == str(block_value)].copy()

def filter_exact_week(df, selected_week):
    if "week_date" not in df.columns:
        return df.copy()
    temp = df.copy()
    temp["week_date"] = pd.to_datetime(temp["week_date"], errors="coerce")
    temp = temp.dropna(subset=["week_date"])
    return temp[temp["week_date"] == selected_week].copy()

def filter_range(df, start_date, end_date):
    if "week_date" not in df.columns:
        return df.copy()
    temp = df.copy()
    temp["week_date"] = pd.to_datetime(temp["week_date"], errors="coerce")
    temp = temp.dropna(subset=["week_date"])
    return temp[(temp["week_date"] >= start_date) & (temp["week_date"] <= end_date)].copy()