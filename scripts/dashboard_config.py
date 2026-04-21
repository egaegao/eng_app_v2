from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

REQUIRED_SHEETS = [
    "weekly_kpi",
    "fuel_ratio",
    "issue_log",
    "action_tracker",
    "inventory",
    "hauling_review",
    "findings_summary",
    "unit_performance",
]

LOWER_IS_BETTER = ["rain_hours", "slippery_hours", "fuel_ratio_mining", "stripping_ratio"]
PRIMARY_METRICS = ["overburden", "coal_getting", "rain_hours", "ewh_ob"]

PLAN_COLOR = "#B8C5D6"
ACTUAL_COLOR = "#1F77B4"
GRID_COLOR = "rgba(0,0,0,0.08)"
BG_COLOR = "white"