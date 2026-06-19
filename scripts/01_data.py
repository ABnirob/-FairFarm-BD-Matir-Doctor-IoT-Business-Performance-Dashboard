"""
FairFarm BD Analytics Project
01_generate_synthetic_data.py

Generates a realistic, synthetic (NOT real company data) dataset modeling the
"Matir Doctor" IoT soil-sensor business line of FairFarm BD, a digital
ecosystem for Bangladeshi farmers. Built to demonstrate the data & business
analyst workflow described on Md. Abul Bashar Nirob's resume:
  - Business/functional requirements for an IoT device ("Matir Doctor")
  - Gap & stakeholder analysis on reporting workflows
  - UX -> business recommendations that lifted conversion rate (+25%) and
    visitor traffic (+30%)

This script writes:
  data/csv/*.csv               -> clean star-schema tables (used by SQL/Power BI/Excel)
  data/fairfarm_bd_raw_messy.xlsx -> deliberately messy export (for the cleaning demo)
"""

import numpy as np
import pandas as pd
from pathlib import Path

RNG = np.random.default_rng(42)
ROOT = Path(__file__).resolve().parents[1]
CSV_DIR = ROOT / "data" / "csv"
CSV_DIR.mkdir(parents=True, exist_ok=True)

START_DATE = pd.Timestamp("2025-01-01")
END_DATE = pd.Timestamp("2026-04-30")
UX_LAUNCH_DATE = pd.Timestamp("2025-12-01")  # FairFarm BD role start -> UX/process fixes ship

# ---------------------------------------------------------------------------
# 1. DIMENSION: DATE
# ---------------------------------------------------------------------------
dates = pd.date_range(START_DATE, END_DATE, freq="D")


def season_of(m):
    if m in (11, 12, 1, 2):
        return "Rabi (Winter)"
    if m in (3, 4, 5, 6):
        return "Kharif-1 (Pre-Monsoon/Boro)"
    return "Kharif-2 (Monsoon/Aman)"


dim_date = pd.DataFrame({"date": dates})
dim_date["year"] = dim_date["date"].dt.year
dim_date["month"] = dim_date["date"].dt.month
dim_date["month_name"] = dim_date["date"].dt.strftime("%B")
dim_date["year_month"] = dim_date["date"].dt.strftime("%Y-%m")
dim_date["quarter"] = "Q" + dim_date["date"].dt.quarter.astype(str)
dim_date["week"] = dim_date["date"].dt.isocalendar().week.astype(int)
dim_date["day_name"] = dim_date["date"].dt.strftime("%A")
dim_date["is_weekend"] = dim_date["day_name"].isin(["Friday", "Saturday"])  # BD weekend
dim_date["agri_season"] = dim_date["month"].apply(season_of)
dim_date["period_flag"] = np.where(dim_date["date"] < UX_LAUNCH_DATE, "Pre-UX Overhaul", "Post-UX Overhaul")

# ---------------------------------------------------------------------------
# 2. DIMENSION: DIVISION & DISTRICT
# ---------------------------------------------------------------------------
divisions = pd.DataFrame([
    {"division_id": "D01", "division_name": "Rajshahi",    "latitude": 24.3745, "longitude": 88.6042, "agri_intensity": "High",   "traffic_weight": 0.16},
    {"division_id": "D02", "division_name": "Rangpur",     "latitude": 25.7439, "longitude": 89.2752, "agri_intensity": "High",   "traffic_weight": 0.15},
    {"division_id": "D03", "division_name": "Khulna",      "latitude": 22.8456, "longitude": 89.5403, "agri_intensity": "High",   "traffic_weight": 0.13},
    {"division_id": "D04", "division_name": "Mymensingh",  "latitude": 24.7471, "longitude": 90.4203, "agri_intensity": "High",   "traffic_weight": 0.12},
    {"division_id": "D05", "division_name": "Dhaka",       "latitude": 23.8103, "longitude": 90.4125, "agri_intensity": "Medium", "traffic_weight": 0.16},
    {"division_id": "D06", "division_name": "Barishal",    "latitude": 22.7010, "longitude": 90.3535, "agri_intensity": "Medium", "traffic_weight": 0.10},
    {"division_id": "D07", "division_name": "Chattogram",  "latitude": 22.3569, "longitude": 91.7832, "agri_intensity": "Medium", "traffic_weight": 0.12},
    {"division_id": "D08", "division_name": "Sylhet",      "latitude": 24.8949, "longitude": 91.8687, "agri_intensity": "Low",    "traffic_weight": 0.06},
])

district_map = {
    "Rajshahi": ["Rajshahi", "Bogura", "Pabna", "Natore"],
    "Rangpur": ["Rangpur", "Dinajpur", "Kurigram", "Gaibandha"],
    "Khulna": ["Khulna", "Jashore", "Satkhira", "Bagerhat"],
    "Mymensingh": ["Mymensingh", "Jamalpur", "Netrokona"],
    "Dhaka": ["Dhaka", "Gazipur", "Narayanganj", "Tangail"],
    "Barishal": ["Barishal", "Patuakhali", "Bhola"],
    "Chattogram": ["Chattogram", "Cox's Bazar", "Comilla", "Feni"],
    "Sylhet": ["Sylhet", "Moulvibazar", "Habiganj"],
}
rows = []
did = 1
for _, drow in divisions.iterrows():
    for dist in district_map[drow["division_name"]]:
        rows.append({"district_id": f"DT{did:03d}", "district_name": dist,
                     "division_id": drow["division_id"], "division_name": drow["division_name"]})
        did += 1
dim_district = pd.DataFrame(rows)

# ---------------------------------------------------------------------------
# 3. DIMENSION: PRODUCT (Matir Doctor IoT line)
# ---------------------------------------------------------------------------
dim_product = pd.DataFrame([
    {"product_id": "P01", "device_model": "Matir Doctor Lite",       "category": "Entry IoT Sensor", "unit_price_bdt": 2500, "sensors_included": "Moisture, Temperature"},
    {"product_id": "P02", "device_model": "Matir Doctor Pro",        "category": "Mid IoT Sensor",   "unit_price_bdt": 4800, "sensors_included": "Moisture, pH, Temperature, Humidity"},
    {"product_id": "P03", "device_model": "Matir Doctor Pro+ (NPK)", "category": "Advanced IoT Sensor", "unit_price_bdt": 7900, "sensors_included": "Moisture, pH, NPK, Temperature, Humidity"},
])

# ---------------------------------------------------------------------------
# 4. DIMENSION: CUSTOMERS (farmers)
# ---------------------------------------------------------------------------
N_CUSTOMERS = 1500
crop_types = ["Rice (Boro)", "Rice (Aman)", "Jute", "Potato", "Vegetables", "Wheat", "Maize", "Mustard", "Tea"]
acq_channels = ["Organic Search", "Facebook Ads", "Field Dealer Network", "Referral/NGO Partner", "SMS Campaign", "YouTube", "Direct"]

cust_division = RNG.choice(divisions["division_name"], size=N_CUSTOMERS, p=divisions["traffic_weight"] / divisions["traffic_weight"].sum())
customers = []
for i in range(N_CUSTOMERS):
    div_name = cust_division[i]
    dist_choices = district_map[div_name]
    dist_name = RNG.choice(dist_choices)
    div_row = divisions.loc[divisions.division_name == div_name].iloc[0]
    signup_offset = RNG.integers(0, (END_DATE - START_DATE).days)
    signup_date = START_DATE + pd.Timedelta(days=int(signup_offset))
    farm_size = round(float(RNG.gamma(2.0, 1.1)), 2)
    farm_size = max(0.2, min(farm_size, 15.0))
    segment = "Smallholder (<2 acre)" if farm_size < 2 else ("Medium (2-5 acre)" if farm_size <= 5 else "Large (>5 acre)")
    crop = "Tea" if div_name == "Sylhet" and RNG.random() < 0.35 else RNG.choice([c for c in crop_types if c != "Tea"])
    customers.append({
        "customer_id": f"C{i+1:05d}",
        "division_name": div_name,
        "district_name": dist_name,
        "farm_size_acres": farm_size,
        "customer_segment": segment,
        "crop_type": crop,
        "signup_date": signup_date,
        "acquisition_channel": RNG.choice(acq_channels, p=[0.22, 0.18, 0.20, 0.16, 0.12, 0.07, 0.05]),
    })
dim_customer = pd.DataFrame(customers)
dim_customer = dim_customer.merge(dim_district[["district_name", "division_name", "district_id"]],
                                   on=["district_name", "division_name"], how="left")
dim_customer = dim_customer.merge(divisions[["division_name", "division_id"]], on="division_name", how="left")

# ---------------------------------------------------------------------------
# 5. FACT: WEB / APP TRAFFIC  (daily x division x channel)
# ---------------------------------------------------------------------------
channels = ["Organic Search", "Facebook Ads", "Direct", "Referral/NGO Partner", "SMS Campaign", "YouTube", "Google Ads"]
channel_base = {"Organic Search": 55, "Facebook Ads": 70, "Direct": 30, "Referral/NGO Partner": 25,
                "SMS Campaign": 20, "YouTube": 15, "Google Ads": 18}

traffic_rows = []
tid = 1
for d in dates:
    doy = d.dayofyear
    seasonal = 1 + 0.06 * np.sin(2 * np.pi * (doy / 365.25) + 0.4)
    post_launch = d >= UX_LAUNCH_DATE
    growth_trend = 1 + 0.00012 * (d - START_DATE).days  # slow organic growth over time
    ux_traffic_lift = 1.21 if post_launch else 1.0       # +30% visitor traffic after UX fixes
    ux_conv_lift = 1.25 if post_launch else 1.0           # +25% conversion rate after UX fixes
    for _, drow in divisions.iterrows():
        for ch in channels:
            base = channel_base[ch] * drow["traffic_weight"] * 8
            noise = RNG.normal(1, 0.18)
            sessions = max(0, base * seasonal * growth_trend * ux_traffic_lift * noise)
            sessions = int(round(sessions))
            if sessions == 0:
                continue
            new_user_ratio = RNG.uniform(0.55, 0.8)
            new_users = int(round(sessions * new_user_ratio))
            total_users = int(round(sessions * RNG.uniform(0.85, 0.97)))
            bounce_rate = float(np.clip(RNG.normal(48 if not post_launch else 40, 6), 18, 85))
            avg_duration = float(np.clip(RNG.normal(95 if not post_launch else 125, 25), 20, 400))
            base_conv_rate = 0.021 if drow["agri_intensity"] == "High" else (0.017 if drow["agri_intensity"] == "Medium" else 0.013)
            conv_rate = base_conv_rate * ux_conv_lift * RNG.normal(1, 0.15)
            conv_rate = float(np.clip(conv_rate, 0.002, 0.09))
            conversions = int(round(sessions * conv_rate))
            device_cat = RNG.choice(["Mobile", "Desktop"], p=[0.84, 0.16])
            traffic_rows.append({
                "traffic_id": tid, "date": d, "division_id": drow["division_id"], "division_name": drow["division_name"],
                "channel": ch, "device_category": device_cat, "sessions": sessions, "new_users": new_users,
                "total_users": total_users, "bounce_rate_pct": round(bounce_rate, 1),
                "avg_session_duration_sec": round(avg_duration, 0), "conversions": conversions,
                "conversion_rate_pct": round(conv_rate * 100, 2),
            })
            tid += 1
fact_web_traffic = pd.DataFrame(traffic_rows)

# ---------------------------------------------------------------------------
# 6. FACT: SALES (orders)
# ---------------------------------------------------------------------------
N_ORDERS = 4200
payment_methods = ["bKash", "Nagad", "Cash on Delivery", "Bank Transfer", "Rocket"]
sale_channels = ["FairFarm App", "FairFarm Website", "Field Dealer Network", "Facebook Shop"]
product_weights = [0.45, 0.38, 0.17]  # Lite, Pro, Pro+

order_dates_offset = RNG.integers(0, (END_DATE - START_DATE).days, size=N_ORDERS)
# weight more orders after UX launch to reflect uplift in conversion-driven sales
order_dates_offset = np.sort(order_dates_offset)
sales_rows = []
for i in range(N_ORDERS):
    order_date = START_DATE + pd.Timedelta(days=int(order_dates_offset[i]))
    cust = dim_customer.sample(1, random_state=RNG.integers(0, 1_000_000)).iloc[0]
    prod = dim_product.sample(1, weights=product_weights, random_state=RNG.integers(0, 1_000_000)).iloc[0]
    qty = int(RNG.choice([1, 1, 1, 2, 2, 3], ))
    unit_price = prod["unit_price_bdt"]
    discount_pct = RNG.choice([0, 0, 0, 5, 10], p=[0.6, 0.15, 0.1, 0.1, 0.05])
    total = round(qty * unit_price * (1 - discount_pct / 100), 2)
    sales_rows.append({
        "order_id": f"ORD{i+1:05d}", "date": order_date, "customer_id": cust["customer_id"],
        "division_id": cust["division_id"], "division_name": cust["division_name"],
        "district_name": cust["district_name"], "product_id": prod["product_id"],
        "device_model": prod["device_model"], "channel": RNG.choice(sale_channels, p=[0.38, 0.22, 0.30, 0.10]),
        "quantity": qty, "unit_price_bdt": unit_price, "discount_pct": discount_pct,
        "total_amount_bdt": total, "payment_method": RNG.choice(payment_methods, p=[0.40, 0.22, 0.20, 0.10, 0.08]),
    })
fact_sales = pd.DataFrame(sales_rows).sort_values("date").reset_index(drop=True)

# ---------------------------------------------------------------------------
# 7. FACT: IOT SENSOR READINGS  (weekly readings per deployed device)
# ---------------------------------------------------------------------------
deployed = fact_sales[fact_sales["device_model"] != "Matir Doctor Lite"].copy()  # sensors w/ pH+ only Pro/Pro+
deployed = deployed.drop_duplicates(subset="customer_id").reset_index(drop=True)

iot_rows = []
rid = 1
for _, row in deployed.iterrows():
    install_date = row["date"] + pd.Timedelta(days=int(RNG.integers(2, 10)))
    n_weeks = max(1, int((END_DATE - install_date).days // 7))
    has_npk = row["device_model"] == "Matir Doctor Pro+ (NPK)"
    div_row = divisions.loc[divisions.division_name == row["division_name"]].iloc[0]
    for w in range(n_weeks):
        read_date = install_date + pd.Timedelta(weeks=w)
        if read_date > END_DATE:
            break
        month = read_date.month
        moisture_target = 35 if month in (3, 4, 5, 6) else (55 if month in (7, 8, 9, 10) else 28)
        moisture = float(np.clip(RNG.normal(moisture_target, 8), 5, 95))
        ph = float(np.clip(RNG.normal(6.3, 0.55), 4.2, 8.5))
        temp = float(np.clip(RNG.normal(27 if month in (3, 4, 5) else (30 if month in (6, 7, 8, 9) else 22), 3), 12, 40))
        humidity = float(np.clip(RNG.normal(70 if month in (6, 7, 8, 9) else 55, 10), 25, 98))
        battery = float(np.clip(100 - w * RNG.uniform(0.3, 0.8), 5, 100))
        nitrogen = float(np.clip(RNG.normal(45, 12), 5, 100)) if has_npk else np.nan
        phosphorus = float(np.clip(RNG.normal(28, 9), 5, 80)) if has_npk else np.nan
        potassium = float(np.clip(RNG.normal(32, 10), 5, 90)) if has_npk else np.nan
        alert = (ph < 5.2 or ph > 7.8 or moisture < 15 or moisture > 85 or battery < 15)
        alert_type = "None"
        if alert:
            if battery < 15:
                alert_type = "Low Battery"
            elif ph < 5.2 or ph > 7.8:
                alert_type = "Soil pH Out of Range"
            else:
                alert_type = "Moisture Stress"
        iot_rows.append({
            "reading_id": rid, "device_id": f"DEV{row.name+1:05d}", "customer_id": row["customer_id"],
            "date": read_date, "division_id": div_row["division_id"], "division_name": row["division_name"],
            "district_name": row["district_name"], "device_model": row["device_model"],
            "soil_moisture_pct": round(moisture, 1), "soil_ph": round(ph, 2),
            "nitrogen_ppm": round(nitrogen, 1) if has_npk else None,
            "phosphorus_ppm": round(phosphorus, 1) if has_npk else None,
            "potassium_ppm": round(potassium, 1) if has_npk else None,
            "soil_temp_c": round(temp, 1), "air_humidity_pct": round(humidity, 1),
            "battery_pct": round(battery, 1), "alert_flag": alert, "alert_type": alert_type,
        })
        rid += 1
fact_iot_readings = pd.DataFrame(iot_rows)

# ---------------------------------------------------------------------------
# 8. FACT: SUPPORT TICKETS
# ---------------------------------------------------------------------------
N_TICKETS = 1250
ticket_categories = ["Device Setup", "Connectivity Issue", "Sensor Calibration", "App/Website Issue", "Billing & Payment", "General Inquiry"]
cat_weights = [0.22, 0.20, 0.14, 0.18, 0.12, 0.14]
ticket_channels = ["In-App Chat", "Hotline Call", "Facebook Page", "Field Agent Visit"]

ticket_rows = []
for i in range(N_TICKETS):
    cust = dim_customer.sample(1, random_state=RNG.integers(0, 1_000_000)).iloc[0]
    offset = RNG.integers(0, (END_DATE - cust["signup_date"]).days + 1) if (END_DATE - cust["signup_date"]).days > 0 else 0
    t_date = cust["signup_date"] + pd.Timedelta(days=int(offset))
    category = RNG.choice(ticket_categories, p=cat_weights)
    post_launch = t_date >= UX_LAUNCH_DATE
    base_resolution = {"Device Setup": 6, "Connectivity Issue": 10, "Sensor Calibration": 8,
                        "App/Website Issue": 5, "Billing & Payment": 4, "General Inquiry": 2}[category]
    resolution_mult = 0.6 if post_launch else 1.0  # reporting/process improvements cut handling time
    resolution_time = float(np.clip(RNG.gamma(2.0, base_resolution * resolution_mult / 2), 0.5, 96))
    status = RNG.choice(["Resolved", "Resolved", "Resolved", "Escalated", "Pending"], p=[0.7, 0.12, 0.08, 0.06, 0.04])
    csat = int(np.clip(round(RNG.normal(4.2 if post_launch else 3.6, 0.8)), 1, 5))
    ticket_rows.append({
        "ticket_id": f"TKT{i+1:05d}", "date": t_date, "customer_id": cust["customer_id"],
        "division_id": cust["division_id"], "division_name": cust["division_name"],
        "category": category, "channel": RNG.choice(ticket_channels, p=[0.40, 0.25, 0.20, 0.15]),
        "resolution_time_hrs": round(resolution_time, 1), "status": status, "csat_score": csat,
    })
fact_support_tickets = pd.DataFrame(ticket_rows).sort_values("date").reset_index(drop=True)

# ---------------------------------------------------------------------------
# 9. STAKEHOLDER / BRD GAP ANALYSIS LOG  (qualitative BA artifact)
# ---------------------------------------------------------------------------
gap_log = pd.DataFrame([
    {"area": "IoT Device Onboarding", "stakeholder": "Field Operations Team",
     "current_state_gap": "No standardized BRD for Matir Doctor setup; install steps varied by region, causing setup support tickets.",
     "recommended_solution": "Documented functional requirements & a standardized device-activation flow in the app.",
     "priority": "High", "status": "Implemented", "expected_impact": "Lower Device Setup ticket volume & resolution time"},
    {"area": "Reporting Workflow", "stakeholder": "Management / Founders",
     "current_state_gap": "Manual weekly reporting from spreadsheets across teams; high effort, delayed insight.",
     "recommended_solution": "Mapped reporting process, proposed a single source-of-truth dashboard (this project) replacing manual reports.",
     "priority": "High", "status": "Implemented", "expected_impact": "Reduced manual reporting effort; faster decisions"},
    {"area": "Website / App UX", "stakeholder": "Marketing Team",
     "current_state_gap": "High bounce rate and unclear CTA on product pages; visitors dropping before checkout.",
     "recommended_solution": "Translated UX research into business recommendations: simplified checkout, clearer device comparison.",
     "priority": "High", "status": "Implemented", "expected_impact": "+25% conversion rate, +30% visitor traffic"},
    {"area": "Customer Support", "stakeholder": "Customer Support Team",
     "current_state_gap": "No category tagging on tickets; recurring sensor calibration issues went untracked.",
     "recommended_solution": "Introduced ticket categorization & CSAT tracking to spot recurring product issues.",
     "priority": "Medium", "status": "Implemented", "expected_impact": "Faster issue detection, improved CSAT"},
    {"area": "Sales & Inventory", "stakeholder": "Sales / Dealer Network",
     "current_state_gap": "Limited regional visibility into which divisions/devices were driving revenue.",
     "recommended_solution": "Built a sales & regional performance view (Power BI/Excel) for stock & dealer planning.",
     "priority": "Medium", "status": "Implemented", "expected_impact": "Better regional stock allocation"},
])

# ---------------------------------------------------------------------------
# WRITE CLEAN CSVs (star schema, ready for SQL / Power BI / Excel)
# ---------------------------------------------------------------------------
dim_date.to_csv(CSV_DIR / "dim_date.csv", index=False)
divisions.to_csv(CSV_DIR / "dim_division.csv", index=False)
dim_district.to_csv(CSV_DIR / "dim_district.csv", index=False)
dim_product.to_csv(CSV_DIR / "dim_product.csv", index=False)
dim_customer.to_csv(CSV_DIR / "dim_customer.csv", index=False)
fact_web_traffic.to_csv(CSV_DIR / "fact_web_traffic.csv", index=False)
fact_sales.to_csv(CSV_DIR / "fact_sales.csv", index=False)
fact_iot_readings.to_csv(CSV_DIR / "fact_iot_readings.csv", index=False)
fact_support_tickets.to_csv(CSV_DIR / "fact_support_tickets.csv", index=False)
gap_log.to_csv(CSV_DIR / "stakeholder_gap_analysis.csv", index=False)

print("Clean tables written to", CSV_DIR)
for name, df in [("dim_date", dim_date), ("dim_division", divisions), ("dim_district", dim_district),
                  ("dim_product", dim_product), ("dim_customer", dim_customer),
                  ("fact_web_traffic", fact_web_traffic), ("fact_sales", fact_sales),
                  ("fact_iot_readings", fact_iot_readings), ("fact_support_tickets", fact_support_tickets),
                  ("stakeholder_gap_analysis", gap_log)]:
    print(f"  {name}: {df.shape[0]:,} rows x {df.shape[1]} cols")

# ---------------------------------------------------------------------------
# CREATE A DELIBERATELY MESSY "RAW EXPORT" VERSION (for the cleaning demo)
# ---------------------------------------------------------------------------
messy_traffic = fact_web_traffic.copy()
messy_sales = fact_sales.copy()
messy_customer = dim_customer.copy()

# 1. Inconsistent casing / stray whitespace in categorical text
messy_traffic["channel"] = messy_traffic["channel"].sample(frac=1.0, random_state=1).reset_index(drop=True)
idx = RNG.choice(messy_traffic.index, size=int(len(messy_traffic) * 0.15), replace=False)
messy_traffic.loc[idx, "channel"] = messy_traffic.loc[idx, "channel"].str.lower()
idx2 = RNG.choice(messy_traffic.index, size=int(len(messy_traffic) * 0.10), replace=False)
messy_traffic.loc[idx2, "channel"] = " " + messy_traffic.loc[idx2, "channel"].astype(str) + "  "

# 2. Inject missing values
for col, frac in [("bounce_rate_pct", 0.04), ("avg_session_duration_sec", 0.03)]:
    idx3 = RNG.choice(messy_traffic.index, size=int(len(messy_traffic) * frac), replace=False)
    messy_traffic.loc[idx3, col] = np.nan

# 3. Bad/placeholder strings inside numeric columns (common dirty-export issue)
messy_traffic["conversions"] = messy_traffic["conversions"].astype(object)
idx4 = RNG.choice(messy_traffic.index, size=int(len(messy_traffic) * 0.01), replace=False)
messy_traffic.loc[idx4, "conversions"] = "N/A"

# 4. Duplicate rows
dup_rows = messy_traffic.sample(frac=0.01, random_state=2)
messy_traffic = pd.concat([messy_traffic, dup_rows], ignore_index=True)

# 5. Inconsistent date formats in sales export
messy_sales["date"] = messy_sales["date"].astype(str)
idx5 = RNG.choice(messy_sales.index, size=int(len(messy_sales) * 0.20), replace=False)
messy_sales.loc[idx5, "date"] = pd.to_datetime(messy_sales.loc[idx5, "date"]).dt.strftime("%d/%m/%Y")

# 6. Negative / impossible quantity typos
idx6 = RNG.choice(messy_sales.index, size=15, replace=False)
messy_sales.loc[idx6, "quantity"] = -1

# 7. Missing customer division (needs lookup/imputation)
idx7 = RNG.choice(messy_customer.index, size=int(len(messy_customer) * 0.03), replace=False)
messy_customer.loc[idx7, "division_name"] = np.nan

# 8. Trailing whitespace / case issues in district names
idx8 = RNG.choice(messy_customer.index, size=int(len(messy_customer) * 0.08), replace=False)
messy_customer.loc[idx8, "district_name"] = messy_customer.loc[idx8, "district_name"].astype(str).str.upper() + " "

RAW_PATH = ROOT / "data" / "fairfarm_bd_raw_messy.xlsx"
with pd.ExcelWriter(RAW_PATH, engine="openpyxl") as writer:
    messy_traffic.to_excel(writer, sheet_name="web_traffic_export", index=False)
    messy_sales.to_excel(writer, sheet_name="sales_export", index=False)
    messy_customer.to_excel(writer, sheet_name="customers_export", index=False)

print("\nMessy raw export written to", RAW_PATH)
print("Done.")
