"""
FairFarm BD Analytics Project
02_data_cleaning_eda.py

Takes the deliberately messy raw export (data/fairfarm_bd_raw_messy.xlsx) and
runs it through a documented cleaning pipeline, producing a data-quality
report plus cleaned CSVs. Then runs exploratory data analysis (EDA) on the
full clean star-schema and saves chart images for the README / case study.
"""

import json
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CSV_DIR = ROOT / "data" / "csv"
IMG_DIR = ROOT / "images"
IMG_DIR.mkdir(parents=True, exist_ok=True)

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.edgecolor": "#3A4A40",
    "axes.labelcolor": "#1F2D24",
    "text.color": "#1F2D24",
    "xtick.color": "#1F2D24",
    "ytick.color": "#1F2D24",
    "axes.facecolor": "#FFFFFF",
    "figure.facecolor": "#FFFFFF",
    "axes.grid": True,
    "grid.color": "#E3E8E3",
    "grid.linewidth": 0.6,
})
GREEN = "#2E7D4F"
GOLD = "#D99B2B"
TERRA = "#C1532A"
NAVY = "#1F2D24"
PALETTE = ["#2E7D4F", "#D99B2B", "#C1532A", "#3E6B8C", "#7AAE6B", "#8C5B3E"]

# ===========================================================================
# PART A — DATA CLEANING DEMO (messy raw export -> clean)
# ===========================================================================
RAW_PATH = ROOT / "data" / "fairfarm_bd_raw_messy.xlsx"
raw = pd.read_excel(RAW_PATH, sheet_name=None)
traffic_raw = raw["web_traffic_export"].copy()
sales_raw = raw["sales_export"].copy()
cust_raw = raw["customers_export"].copy()

report = {"web_traffic_export": {}, "sales_export": {}, "customers_export": {}}

# --- clean web_traffic_export -------------------------------------------------
report["web_traffic_export"]["rows_before"] = len(traffic_raw)
traffic_raw["channel"] = traffic_raw["channel"].astype(str).str.strip().str.title()
traffic_raw["channel"] = traffic_raw["channel"].replace({"Sms Campaign": "SMS Campaign"})

n_dupes = traffic_raw.duplicated().sum()
traffic_raw = traffic_raw.drop_duplicates()
report["web_traffic_export"]["duplicate_rows_removed"] = int(n_dupes)

n_bad_conversions = (traffic_raw["conversions"].astype(str) == "N/A").sum()
traffic_raw["conversions"] = pd.to_numeric(traffic_raw["conversions"], errors="coerce")
traffic_raw["conversions"] = traffic_raw["conversions"].fillna(
    (traffic_raw["sessions"] * traffic_raw["conversion_rate_pct"] / 100).round()
)
report["web_traffic_export"]["non_numeric_conversions_fixed"] = int(n_bad_conversions)

for col in ["bounce_rate_pct", "avg_session_duration_sec"]:
    n_missing = traffic_raw[col].isna().sum()
    traffic_raw[col] = traffic_raw.groupby("division_name")[col].transform(lambda s: s.fillna(s.median()))
    report["web_traffic_export"][f"missing_{col}_imputed_by_division_median"] = int(n_missing)

report["web_traffic_export"]["rows_after"] = len(traffic_raw)
traffic_clean = traffic_raw

# --- clean sales_export --------------------------------------------------------
report["sales_export"]["rows_before"] = len(sales_raw)


def parse_mixed_date(s):
    s = str(s)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return pd.to_datetime(s, format=fmt)
        except ValueError:
            continue
    return pd.to_datetime(s, errors="coerce")


n_alt_format = 0
parsed_dates = []
for v in sales_raw["date"]:
    if "/" in str(v):
        n_alt_format += 1
    parsed_dates.append(parse_mixed_date(v))
sales_raw["date"] = parsed_dates
report["sales_export"]["mixed_date_formats_standardized"] = int(n_alt_format)

n_bad_qty = (sales_raw["quantity"] < 0).sum()
sales_raw.loc[sales_raw["quantity"] < 0, "quantity"] = np.nan
sales_raw["quantity"] = sales_raw["quantity"].fillna(sales_raw["quantity"].median())
report["sales_export"]["invalid_negative_quantity_fixed"] = int(n_bad_qty)
report["sales_export"]["rows_after"] = len(sales_raw)
sales_clean = sales_raw

# --- clean customers_export ----------------------------------------------------
report["customers_export"]["rows_before"] = len(cust_raw)
cust_raw["district_name"] = cust_raw["district_name"].astype(str).str.strip().str.title()

n_missing_division = cust_raw["division_name"].isna().sum()
district_to_division = pd.read_csv(CSV_DIR / "dim_district.csv")[["district_name", "division_name"]] \
    .drop_duplicates().set_index("district_name")["division_name"]
cust_raw["division_name"] = cust_raw.apply(
    lambda r: district_to_division.get(r["district_name"], r["division_name"])
    if pd.isna(r["division_name"]) else r["division_name"], axis=1
)
report["customers_export"]["missing_division_imputed_via_district_lookup"] = int(n_missing_division)
report["customers_export"]["rows_after"] = len(cust_raw)
cust_clean = cust_raw

with open(ROOT / "docs" / "data_quality_report.json", "w") as f:
    json.dump(report, f, indent=2, default=str)

print("=== DATA QUALITY / CLEANING REPORT ===")
print(json.dumps(report, indent=2, default=str))

# ===========================================================================
# PART B — EDA on the full clean star schema
# ===========================================================================
dim_date = pd.read_csv(CSV_DIR / "dim_date.csv", parse_dates=["date"])
fact_web_traffic = pd.read_csv(CSV_DIR / "fact_web_traffic.csv", parse_dates=["date"])
fact_sales = pd.read_csv(CSV_DIR / "fact_sales.csv", parse_dates=["date"])
fact_iot = pd.read_csv(CSV_DIR / "fact_iot_readings.csv", parse_dates=["date"])
fact_tickets = pd.read_csv(CSV_DIR / "fact_support_tickets.csv", parse_dates=["date"])
dim_customer = pd.read_csv(CSV_DIR / "dim_customer.csv", parse_dates=["signup_date"])

traffic = fact_web_traffic.merge(dim_date[["date", "period_flag", "year_month"]], on="date", how="left")

# --- 1. Before/after UX overhaul: conversion rate & sessions -----------------
agg = traffic.groupby("period_flag").agg(
    avg_conversion_rate=("conversion_rate_pct", "mean"),
    avg_daily_sessions=("sessions", lambda s: s.sum() / traffic.loc[s.index, "date"].nunique()),
).round(2)
pre = agg.loc["Pre-UX Overhaul"]
post = agg.loc["Post-UX Overhaul"]
conv_lift_pct = round((post["avg_conversion_rate"] / pre["avg_conversion_rate"] - 1) * 100, 1)
traffic_lift_pct = round((post["avg_daily_sessions"] / pre["avg_daily_sessions"] - 1) * 100, 1)

from scipy import stats
daily_conv = traffic.groupby(["date", "period_flag"])["conversion_rate_pct"].mean().reset_index()
pre_vals = daily_conv.loc[daily_conv.period_flag == "Pre-UX Overhaul", "conversion_rate_pct"]
post_vals = daily_conv.loc[daily_conv.period_flag == "Post-UX Overhaul", "conversion_rate_pct"]
tstat, pval = stats.ttest_ind(post_vals, pre_vals, equal_var=False)

monthly = traffic.groupby("year_month").agg(sessions=("sessions", "sum"), conversions=("conversions", "sum")).reset_index()
monthly["conversion_rate_pct"] = (monthly["conversions"] / monthly["sessions"] * 100).round(2)

fig, ax1 = plt.subplots(figsize=(10, 5))
ax2 = ax1.twinx()
ax1.bar(monthly["year_month"], monthly["sessions"], color=GREEN, alpha=0.85, label="Sessions")
ax2.plot(monthly["year_month"], monthly["conversion_rate_pct"], color=TERRA, marker="o", linewidth=2.5, label="Conversion Rate %")
ax1.axvline(x="2025-12", color=NAVY, linestyle="--", linewidth=1.2)
ax1.text(monthly[monthly.year_month == "2025-12"].index[0] if "2025-12" in monthly.year_month.values else 11,
          ax1.get_ylim()[1] * 0.92, "UX Overhaul\nshipped", fontsize=8, color=NAVY)
ax1.set_ylabel("Monthly Sessions")
ax2.set_ylabel("Conversion Rate (%)", color=TERRA)
ax1.set_xticks(range(len(monthly)))
ax1.set_xticklabels(monthly["year_month"], rotation=45, ha="right", fontsize=8)
ax1.set_title("FairFarm BD — Monthly Traffic & Conversion Rate (Before / After UX Overhaul)", fontsize=12, weight="bold")
fig.tight_layout()
fig.savefig(IMG_DIR / "eda_traffic_conversion_trend.png", dpi=160)
plt.close(fig)

# --- 2. Revenue by division ---------------------------------------------------
rev_div = fact_sales.groupby("division_name")["total_amount_bdt"].sum().sort_values(ascending=False)
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.barh(rev_div.index[::-1], rev_div.values[::-1], color=GREEN)
ax.bar_label(bars, labels=[f"BDT {v:,.0f}" for v in rev_div.values[::-1]], padding=4, fontsize=8)
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"BDT {x/1e6:.1f}M"))
ax.set_title("Total Revenue by Division", fontsize=12, weight="bold")
fig.tight_layout()
fig.savefig(IMG_DIR / "eda_revenue_by_division.png", dpi=160)
plt.close(fig)

# --- 3. Device mix & IoT alerts ----------------------------------------------
device_mix = fact_sales.groupby("device_model")["quantity"].sum().sort_values(ascending=False)
alert_mix = fact_iot[fact_iot.alert_flag]["alert_type"].value_counts()

fig, axes = plt.subplots(1, 2, figsize=(11, 5))
axes[0].pie(device_mix.values, labels=device_mix.index, autopct="%1.0f%%", colors=PALETTE,
            wedgeprops=dict(width=0.42, edgecolor="white"))
axes[0].set_title("Devices Sold by Model", fontsize=11, weight="bold")
axes[1].pie(alert_mix.values, labels=alert_mix.index, autopct="%1.0f%%", colors=PALETTE[1:],
            wedgeprops=dict(width=0.42, edgecolor="white"))
axes[1].set_title("IoT Sensor Alerts by Type", fontsize=11, weight="bold")
fig.tight_layout()
fig.savefig(IMG_DIR / "eda_device_alert_mix.png", dpi=160)
plt.close(fig)

# --- 4. Support ticket resolution time by category ----------------------------
res = fact_tickets.groupby("category")["resolution_time_hrs"].mean().sort_values()
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.barh(res.index, res.values, color=GOLD)
ax.bar_label(bars, fmt="%.1f hrs", padding=4, fontsize=8)
ax.set_title("Avg. Support Ticket Resolution Time by Category", fontsize=12, weight="bold")
fig.tight_layout()
fig.savefig(IMG_DIR / "eda_ticket_resolution.png", dpi=160)
plt.close(fig)

# ===========================================================================
# PART C — Save cleaned demo exports + EDA summary
# ===========================================================================
CLEAN_DEMO_DIR = ROOT / "data" / "cleaned_demo"
CLEAN_DEMO_DIR.mkdir(exist_ok=True)
traffic_clean.to_csv(CLEAN_DEMO_DIR / "web_traffic_cleaned.csv", index=False)
sales_clean.to_csv(CLEAN_DEMO_DIR / "sales_cleaned.csv", index=False)
cust_clean.to_csv(CLEAN_DEMO_DIR / "customers_cleaned.csv", index=False)

summary = {
    "pre_ux_avg_conversion_rate_pct": round(float(pre["avg_conversion_rate"]), 2),
    "post_ux_avg_conversion_rate_pct": round(float(post["avg_conversion_rate"]), 2),
    "conversion_rate_lift_pct": conv_lift_pct,
    "pre_ux_avg_daily_sessions": round(float(pre["avg_daily_sessions"]), 1),
    "post_ux_avg_daily_sessions": round(float(post["avg_daily_sessions"]), 1),
    "visitor_traffic_lift_pct": traffic_lift_pct,
    "welch_t_stat": round(float(tstat), 3),
    "p_value": float(pval),
    "statistically_significant_at_0.05": bool(pval < 0.05),
    "total_revenue_bdt": float(fact_sales["total_amount_bdt"].sum()),
    "total_orders": int(len(fact_sales)),
    "total_devices_deployed": int(fact_iot["device_id"].nunique()),
    "pct_iot_readings_with_alert": round(float(fact_iot["alert_flag"].mean() * 100), 2),
    "avg_csat_score": round(float(fact_tickets["csat_score"].mean()), 2),
}
with open(ROOT / "docs" / "eda_kpi_summary.json", "w") as f:
    json.dump(summary, f, indent=2)

print("\n=== EDA KPI SUMMARY ===")
print(json.dumps(summary, indent=2))
print("\nCharts saved to", IMG_DIR)
print("Cleaned demo CSVs saved to", CLEAN_DEMO_DIR)
