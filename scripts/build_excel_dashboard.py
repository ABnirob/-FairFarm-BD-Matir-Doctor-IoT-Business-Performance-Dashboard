"""
FairFarm BD Analytics Project
build_excel_dashboard.py

Builds a polished, formula-driven Excel dashboard:
  excel_dashboard/FairFarm_BD_Excel_Dashboard.xlsx

Sheets:
  Cover            - project / analyst summary
  Dashboard        - KPI cards + charts + division filter
  Calc_Helpers     - formula-driven summary tables that feed the charts
  Data_Sales / Data_Traffic / Data_Tickets / Data_IoT_Monthly / Data_Customers

All KPI and chart-source numbers are live Excel formulas (SUMIF/SUMIFS/
AVERAGEIF/COUNTIF) referencing the raw data sheets, not hardcoded values.
"""
from pathlib import Path
import pandas as pd
import xlsxwriter
from openpyxl.utils import get_column_letter as gcl

ROOT = Path(__file__).resolve().parents[1]
CSV_DIR = ROOT / "data" / "csv"
OUT_PATH = ROOT / "excel_dashboard" / "FairFarm_BD_Excel_Dashboard.xlsx"
OUT_PATH.parent.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Load clean data
# ---------------------------------------------------------------------------
sales = pd.read_csv(CSV_DIR / "fact_sales.csv", parse_dates=["date"])
traffic = pd.read_csv(CSV_DIR / "fact_web_traffic.csv", parse_dates=["date"])
tickets = pd.read_csv(CSV_DIR / "fact_support_tickets.csv", parse_dates=["date"])
iot = pd.read_csv(CSV_DIR / "fact_iot_readings.csv", parse_dates=["date"])
customers = pd.read_csv(CSV_DIR / "dim_customer.csv", parse_dates=["signup_date"])
divisions = pd.read_csv(CSV_DIR / "dim_division.csv")

for df in (sales, traffic, tickets):
    df["year_month"] = df["date"].dt.strftime("%Y-%m")
    df["date"] = df["date"].dt.strftime("%Y-%m-%d")

iot["year_month"] = iot["date"].dt.strftime("%Y-%m")
iot_monthly = iot.groupby(["year_month", "division_name"]).agg(
    avg_soil_moisture_pct=("soil_moisture_pct", "mean"),
    avg_soil_ph=("soil_ph", "mean"),
    avg_battery_pct=("battery_pct", "mean"),
    alert_rate_pct=("alert_flag", "mean"),
    active_devices=("device_id", "nunique"),
).reset_index()
iot_monthly["avg_soil_moisture_pct"] = iot_monthly["avg_soil_moisture_pct"].round(1)
iot_monthly["avg_soil_ph"] = iot_monthly["avg_soil_ph"].round(2)
iot_monthly["avg_battery_pct"] = iot_monthly["avg_battery_pct"].round(1)
iot_monthly["alert_rate_pct"] = (iot_monthly["alert_rate_pct"] * 100).round(2)

customers_out = customers.copy()
customers_out["signup_date"] = customers_out["signup_date"].dt.strftime("%Y-%m-%d")

month_list = sorted(sales["year_month"].unique())
district_rev_order = (sales.groupby(["district_name", "division_name"])["total_amount_bdt"]
                       .sum().sort_values(ascending=False).reset_index())
top_districts = district_rev_order.head(10)[["district_name", "division_name"]]

ticket_categories = sorted(tickets["category"].unique())
device_models = ["Matir Doctor Lite", "Matir Doctor Pro", "Matir Doctor Pro+ (NPK)"]
division_names = list(divisions.sort_values("division_name")["division_name"])

# ---------------------------------------------------------------------------
# Workbook + formats
# ---------------------------------------------------------------------------
wb = xlsxwriter.Workbook(str(OUT_PATH))

FOREST = "#1F3D2C"
GREEN = "#2E7D4F"
LIGHT_GREEN = "#E8F3EC"
GOLD = "#D99B2B"
TERRA = "#C1532A"
CREAM = "#FBF8F1"
SLATE = "#3A4A40"
WHITE = "#FFFFFF"

f_title = wb.add_format({"bold": True, "font_size": 20, "font_color": WHITE, "bg_color": FOREST,
                          "valign": "vcenter", "indent": 1, "font_name": "Carlito"})
f_cover_h = wb.add_format({"bold": True, "font_size": 14, "font_color": FOREST, "font_name": "Carlito"})
f_cover_p = wb.add_format({"font_size": 10.5, "font_color": SLATE, "font_name": "Carlito", "valign": "top",
                            "text_wrap": True})
f_cell = wb.add_format({"font_name": "Carlito", "font_size": 10})
f_cell_num = wb.add_format({"font_name": "Carlito", "font_size": 10, "num_format": "#,##0"})
f_cell_money = wb.add_format({"font_name": "Carlito", "font_size": 10, "num_format": '"BDT "#,##0'})
f_cell_pct = wb.add_format({"font_name": "Carlito", "font_size": 10, "num_format": "0.00\"%\""})

# ---------------------------------------------------------------------------
# COVER SHEET
# ---------------------------------------------------------------------------
cov = wb.add_worksheet("Cover")
cov.hide_gridlines(2)
cov.set_column("A:A", 3)
cov.set_column("B:B", 100)
cov.set_row(0, 70)
cov.merge_range("A1:H1", "FairFarm BD — \"Matir Doctor\" IoT Business Performance Dashboard", f_title)
cov.write("B3", "Data & Business Analyst Portfolio Project", f_cover_h)
cov.write("B5", "Prepared by Md. Abul Bashar Nirob — Data & Business Analyst, FairFarm Bangladesh", f_cover_p)
cov.write("B7",
          "This workbook demonstrates the analyst workflow behind FairFarm BD's Matir Doctor IoT soil-sensor "
          "business: documenting requirements, cleaning & modeling data, and reporting KPIs to stakeholders -- "
          "the same work that delivered a +25% conversion-rate lift and +30% visitor-traffic growth referenced "
          "on the analyst's resume.", f_cover_p)
cov.write("B9", "IMPORTANT -- Synthetic data notice", f_cover_h)
cov.write("B10",
          "All figures in this workbook are generated from a synthetic dataset built to realistically mirror "
          "FairFarm BD's Matir Doctor product line and Bangladesh market context. Real proprietary company "
          "data is not published here; this is a portfolio recreation of the analytical workflow and a "
          "validation of the resume's documented results.", f_cover_p)
cov.write("B12", "How to use this workbook", f_cover_h)
cov.write("B13",
          "1) Go to the 'Dashboard' tab.  2) Use the Division filter (top-right) to drill into a single "
          "division -- KPI cards update instantly.  3) 'Calc_Helpers' holds the formula-driven summary tables "
          "that feed every chart.  4) Raw, cleaned tables live in the 'Data_*' tabs (Excel Tables -- filter/sort "
          "freely).", f_cover_p)
for r in (3, 4, 6, 8, 9, 11, 12, 13):
    cov.set_row(r, 18 if r in (3, 8, 12) else 50)

# ---------------------------------------------------------------------------
# DATA SHEETS  (write as Excel Tables)
# ---------------------------------------------------------------------------

def write_table(ws_name, df, money_cols=(), pct_cols=()):
    ws = wb.add_worksheet(ws_name)
    n_rows, n_cols = df.shape
    cols = list(df.columns)
    for j, c in enumerate(cols):
        try:
            width = max(11, min(28, int(df[c].astype(str).str.len().quantile(0.9)) + 4))
        except Exception:
            width = 14
        ws.set_column(j, j, width)
    table_cols = []
    for c in cols:
        fmt = None
        if c in money_cols:
            fmt = f_cell_money
        elif c in pct_cols:
            fmt = f_cell_pct
        table_cols.append({"header": c.replace("_", " ").title(), "format": fmt})
    ws.add_table(0, 0, n_rows, n_cols - 1, {
        "data": df.values.tolist(),
        "columns": table_cols,
        "style": "Table Style Medium 7",
        "name": ws_name,
    })
    ws.freeze_panes(1, 0)
    return ws

write_table("Data_Sales", sales, money_cols=("unit_price_bdt", "total_amount_bdt"), pct_cols=("discount_pct",))
write_table("Data_Traffic", traffic, pct_cols=("bounce_rate_pct", "conversion_rate_pct"))
write_table("Data_Tickets", tickets)
write_table("Data_IoT_Monthly", iot_monthly, pct_cols=("alert_rate_pct",))
write_table("Data_Customers", customers_out)

# ---------------------------------------------------------------------------
# Column-letter lookups & range helpers for formula building
# ---------------------------------------------------------------------------

def col_letters(df):
    return {c: gcl(i + 1) for i, c in enumerate(df.columns)}

SC, TC, KC, IC = col_letters(sales), col_letters(traffic), col_letters(tickets), col_letters(iot_monthly)
n_sales, n_traffic, n_tickets, n_iot = len(sales), len(traffic), len(tickets), len(iot_monthly)


def rng(sheet, col_map, key, n):
    return f"{sheet}!${col_map[key]}$2:${col_map[key]}${n + 1}"


S_DATE = rng("Data_Sales", SC, "date", n_sales)
S_YM = rng("Data_Sales", SC, "year_month", n_sales)
S_DIV = rng("Data_Sales", SC, "division_name", n_sales)
S_DIST = rng("Data_Sales", SC, "district_name", n_sales)
S_MODEL = rng("Data_Sales", SC, "device_model", n_sales)
S_QTY = rng("Data_Sales", SC, "quantity", n_sales)
S_AMT = rng("Data_Sales", SC, "total_amount_bdt", n_sales)

T_YM = rng("Data_Traffic", TC, "year_month", n_traffic)
T_DIV = rng("Data_Traffic", TC, "division_name", n_traffic)
T_SESS = rng("Data_Traffic", TC, "sessions", n_traffic)
T_CONV = rng("Data_Traffic", TC, "conversions", n_traffic)
T_CH = rng("Data_Traffic", TC, "channel", n_traffic)

K_YM = rng("Data_Tickets", KC, "year_month", n_tickets)
K_DIV = rng("Data_Tickets", KC, "division_name", n_tickets)
K_CAT = rng("Data_Tickets", KC, "category", n_tickets)
K_RES = rng("Data_Tickets", KC, "resolution_time_hrs", n_tickets)
K_CSAT = rng("Data_Tickets", KC, "csat_score", n_tickets)

I_YM = rng("Data_IoT_Monthly", IC, "year_month", n_iot)
I_DIV = rng("Data_IoT_Monthly", IC, "division_name", n_iot)
I_ALERT = rng("Data_IoT_Monthly", IC, "alert_rate_pct", n_iot)
I_MOIST = rng("Data_IoT_Monthly", IC, "avg_soil_moisture_pct", n_iot)
I_DEV = rng("Data_IoT_Monthly", IC, "active_devices", n_iot)

# ===========================================================================
# CALC_HELPERS SHEET — every chart/KPI source table, built with formulas
# ===========================================================================
ch = wb.add_worksheet("Calc_Helpers")
ch.hide_gridlines(2)
ch.set_tab_color(SLATE)
ch.write("A1", "Calc_Helpers — formula-driven summary tables feeding the Dashboard charts. Do not delete.",
          wb.add_format({"italic": True, "font_size": 9, "font_color": "#7C8C81", "font_name": "Carlito"}))

f_h = wb.add_format({"bold": True, "font_color": WHITE, "bg_color": GREEN, "font_name": "Carlito", "border": 1})

# --- Monthly_Trend table (rows 3..) ---
r0 = 2
ch.write_row(r0, 0, ["YearMonth", "Revenue_BDT", "Orders", "Sessions", "Conversions",
                      "ConversionRate_Pct", "ActiveDevices", "AvgCSAT"], f_h)
for i, ym in enumerate(month_list):
    r = r0 + 1 + i
    ch.write(r, 0, ym, f_cell)
    ch.write_formula(r, 1, f'=SUMIF({S_YM},$A{r+1},{S_AMT})', f_cell_money)
    ch.write_formula(r, 2, f'=COUNTIF({S_YM},$A{r+1})', f_cell_num)
    ch.write_formula(r, 3, f'=SUMIF({T_YM},$A{r+1},{T_SESS})', f_cell_num)
    ch.write_formula(r, 4, f'=SUMIF({T_YM},$A{r+1},{T_CONV})', f_cell_num)
    ch.write_formula(r, 5, f'=IF($D{r+1}=0,0,$E{r+1}/$D{r+1}*100)', f_cell_pct)
    ch.write_formula(r, 6, f'=SUMIF({I_YM},$A{r+1},{I_DEV})', f_cell_num)
    ch.write_formula(r, 7, f'=IFERROR(AVERAGEIF({K_YM},$A{r+1},{K_CSAT}),0)', wb.add_format(
        {"font_name": "Carlito", "font_size": 10, "num_format": "0.00"}))
monthly_trend_last_row = r0 + len(month_list)
ch.set_column("A:H", 14)

# --- Division_Summary table ---
r1 = monthly_trend_last_row + 3
ch.write_row(r1, 0, ["Division", "Revenue_BDT", "Orders", "Sessions", "Conversions",
                      "ConversionRate_Pct", "DevicesDeployed", "AlertRate_Pct"], f_h)
for i, dv in enumerate(division_names):
    r = r1 + 1 + i
    ch.write(r, 0, dv, f_cell)
    ch.write_formula(r, 1, f'=SUMIF({S_DIV},$A{r+1},{S_AMT})', f_cell_money)
    ch.write_formula(r, 2, f'=COUNTIF({S_DIV},$A{r+1})', f_cell_num)
    ch.write_formula(r, 3, f'=SUMIF({T_DIV},$A{r+1},{T_SESS})', f_cell_num)
    ch.write_formula(r, 4, f'=SUMIF({T_DIV},$A{r+1},{T_CONV})', f_cell_num)
    ch.write_formula(r, 5, f'=IF($D{r+1}=0,0,$E{r+1}/$D{r+1}*100)', f_cell_pct)
    ch.write_formula(r, 6, f'=SUMIFS({I_DEV},{I_DIV},$A{r+1},{I_YM},"{month_list[-1]}")', f_cell_num)
    ch.write_formula(r, 7, f'=IFERROR(AVERAGEIF({I_DIV},$A{r+1},{I_ALERT}),0)', f_cell_pct)
division_summary_last_row = r1 + len(division_names)

# --- Product_Summary table ---
r2 = division_summary_last_row + 3
ch.write_row(r2, 0, ["DeviceModel", "UnitsSold", "Revenue_BDT"], f_h)
for i, m in enumerate(device_models):
    r = r2 + 1 + i
    ch.write(r, 0, m, f_cell)
    ch.write_formula(r, 1, f'=SUMIF({S_MODEL},$A{r+1},{S_QTY})', f_cell_num)
    ch.write_formula(r, 2, f'=SUMIF({S_MODEL},$A{r+1},{S_AMT})', f_cell_money)
product_summary_last_row = r2 + len(device_models)

# --- Ticket_Summary table ---
r3 = product_summary_last_row + 3
ch.write_row(r3, 0, ["Category", "Tickets", "AvgResolution_Hrs", "AvgCSAT"], f_h)
for i, c in enumerate(ticket_categories):
    r = r3 + 1 + i
    ch.write(r, 0, c, f_cell)
    ch.write_formula(r, 1, f'=COUNTIF({K_CAT},$A{r+1})', f_cell_num)
    ch.write_formula(r, 2, f'=IFERROR(AVERAGEIF({K_CAT},$A{r+1},{K_RES}),0)', wb.add_format(
        {"font_name": "Carlito", "font_size": 10, "num_format": "0.0"}))
    ch.write_formula(r, 3, f'=IFERROR(AVERAGEIF({K_CAT},$A{r+1},{K_CSAT}),0)', wb.add_format(
        {"font_name": "Carlito", "font_size": 10, "num_format": "0.00"}))
ticket_summary_last_row = r3 + len(ticket_categories)

# --- Top_Districts table ---
r4 = ticket_summary_last_row + 3
ch.write_row(r4, 0, ["District", "Division", "Revenue_BDT"], f_h)
for i, row in top_districts.reset_index(drop=True).iterrows():
    r = r4 + 1 + i
    ch.write(r, 0, row["district_name"], f_cell)
    ch.write(r, 1, row["division_name"], f_cell)
    ch.write_formula(r, 2, f'=SUMIFS({S_AMT},{S_DIST},$A{r+1},{S_DIV},$B{r+1})', f_cell_money)
top_districts_last_row = r4 + len(top_districts)

# --- Channel_Summary table (acquisition / traffic channel performance) ---
channels_list = sorted(traffic["channel"].unique())
r5 = top_districts_last_row + 3
ch.write_row(r5, 0, ["Channel", "Sessions", "Conversions", "ConversionRate_Pct"], f_h)
for i, c in enumerate(channels_list):
    r = r5 + 1 + i
    ch.write(r, 0, c, f_cell)
    ch.write_formula(r, 1, f'=SUMIF({T_CH},$A{r+1},{T_SESS})', f_cell_num)
    ch.write_formula(r, 2, f'=SUMIF({T_CH},$A{r+1},{T_CONV})', f_cell_num)
    ch.write_formula(r, 3, f'=IF($B{r+1}=0,0,$C{r+1}/$B{r+1}*100)', f_cell_pct)
channel_summary_last_row = r5 + len(channels_list)

print("Calc_Helpers sheet written.")
print("  Monthly_Trend rows:", r0 + 1, "-", monthly_trend_last_row)
print("  Division_Summary rows:", r1 + 1, "-", division_summary_last_row)
print("  Product_Summary rows:", r2 + 1, "-", product_summary_last_row)
print("  Ticket_Summary rows:", r3 + 1, "-", ticket_summary_last_row)
print("  Top_Districts rows:", r4 + 1, "-", top_districts_last_row)
print("  Channel_Summary rows:", r5 + 1, "-", channel_summary_last_row)

# --- Excel (1-based) row bookkeeping for Dashboard chart wiring --------------
MT_HDR, MT_FIRST, MT_LAST = r0 + 1, r0 + 2, monthly_trend_last_row + 1
DS_HDR, DS_FIRST, DS_LAST = r1 + 1, r1 + 2, division_summary_last_row + 1
PS_HDR, PS_FIRST, PS_LAST = r2 + 1, r2 + 2, product_summary_last_row + 1
TS_HDR, TS_FIRST, TS_LAST = r3 + 1, r3 + 2, ticket_summary_last_row + 1
TD_HDR, TD_FIRST, TD_LAST = r4 + 1, r4 + 2, top_districts_last_row + 1
CS_HDR, CS_FIRST, CS_LAST = r5 + 1, r5 + 2, channel_summary_last_row + 1

# ===========================================================================
# ADDITIONAL FORMATS FOR DASHBOARD
# ===========================================================================
f_section = wb.add_format({"bold": True, "font_size": 12, "font_color": WHITE, "bg_color": GREEN,
                            "valign": "vcenter", "indent": 1, "font_name": "Carlito"})
f_kpi_label = wb.add_format({"font_size": 9.5, "font_color": "#5C6F63", "bg_color": WHITE, "valign": "top",
                              "align": "left", "indent": 1, "font_name": "Carlito", "top": 1, "left": 1,
                              "right": 1, "top_color": "#D7E4DB", "left_color": "#D7E4DB",
                              "right_color": "#D7E4DB"})
f_kpi_value_num = wb.add_format({"bold": True, "font_size": 19, "font_color": FOREST, "bg_color": WHITE,
                                  "valign": "bottom", "align": "left", "indent": 1, "font_name": "Carlito",
                                  "bottom": 1, "left": 1, "right": 1, "bottom_color": "#D7E4DB",
                                  "left_color": "#D7E4DB", "right_color": "#D7E4DB", "num_format": "#,##0"})
f_kpi_value_pct = wb.add_format({"bold": True, "font_size": 19, "font_color": GOLD, "bg_color": WHITE,
                                  "valign": "bottom", "align": "left", "indent": 1, "font_name": "Carlito",
                                  "bottom": 1, "left": 1, "right": 1, "bottom_color": "#D7E4DB",
                                  "left_color": "#D7E4DB", "right_color": "#D7E4DB", "num_format": "0.00\"%\""})
f_kpi_value_money = wb.add_format({"bold": True, "font_size": 16, "font_color": FOREST, "bg_color": WHITE,
                                    "valign": "bottom", "align": "left", "indent": 1, "font_name": "Carlito",
                                    "bottom": 1, "left": 1, "right": 1, "bottom_color": "#D7E4DB",
                                    "left_color": "#D7E4DB", "right_color": "#D7E4DB",
                                    "num_format": '"BDT "#,##0'})
f_kpi_value_dec = wb.add_format({"bold": True, "font_size": 19, "font_color": TERRA, "bg_color": WHITE,
                                  "valign": "bottom", "align": "left", "indent": 1, "font_name": "Carlito",
                                  "bottom": 1, "left": 1, "right": 1, "bottom_color": "#D7E4DB",
                                  "left_color": "#D7E4DB", "right_color": "#D7E4DB", "num_format": "0.00"})
f_kpi_accent = wb.add_format({"bg_color": GOLD})
f_filter_label = wb.add_format({"bold": True, "font_size": 10, "font_color": WHITE, "bg_color": FOREST,
                                 "align": "right", "valign": "vcenter", "font_name": "Carlito", "indent": 1})
f_filter_cell = wb.add_format({"bold": True, "font_size": 11, "font_color": WHITE, "bg_color": GOLD,
                                "align": "center", "valign": "vcenter", "font_name": "Carlito", "border": 1})
f_note = wb.add_format({"italic": True, "font_size": 8.5, "font_color": "#7C8C81", "font_name": "Carlito"})
f_chart_caption = wb.add_format({"bold": True, "font_size": 10.5, "font_color": FOREST, "font_name": "Carlito"})

CHART_FONT = {"name": "Carlito", "size": 9}
TITLE_FONT = {"name": "Carlito", "size": 11, "bold": True, "color": FOREST}

# ===========================================================================
# DASHBOARD SHEET
# ===========================================================================
dash = wb.add_worksheet("Dashboard")
dash.hide_gridlines(2)
dash.set_column("A:A", 2.5)
for col, w in zip("BCDEFGHIJKLMNOPQ", [16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16]):
    dash.set_column(f"{col}:{col}", w)

dash.set_row(0, 56)
dash.merge_range("A1:Q1", "  FairFarm BD  —  Matir Doctor IoT Business Performance Dashboard", f_title)

dash.set_row(1, 22)
dash.merge_range("A2:Q2",
                  "  Data & Business Analyst Portfolio Project · Md. Abul Bashar Nirob · Synthetic dataset modeled on FairFarm BD's IoT soil-sensor business",
                  wb.add_format({"italic": True, "font_size": 10, "font_color": WHITE, "bg_color": SLATE,
                                  "valign": "vcenter", "indent": 1, "font_name": "Carlito"}))

dash.write("L4", "Filter by Division:", f_filter_label)
dash.write("M4", "All", f_filter_cell)
dash.data_validation("M4", {"validate": "list", "source": ["All"] + division_names})
dash.write("N4", "", wb.add_format({"bg_color": FOREST}))

FILTER = "$M$4"

kpis = [
    ("TOTAL REVENUE", f'=IF({FILTER}="All",SUM({S_AMT}),SUMIF({S_DIV},{FILTER},{S_AMT}))', f_kpi_value_money),
    ("DEVICES SOLD (UNITS)", f'=IF({FILTER}="All",SUM({S_QTY}),SUMIF({S_DIV},{FILTER},{S_QTY}))', f_kpi_value_num),
    ("AVG CONVERSION RATE",
     f'=IF({FILTER}="All",SUM({T_CONV})/SUM({T_SESS})*100,SUMIF({T_DIV},{FILTER},{T_CONV})/SUMIF({T_DIV},{FILTER},{T_SESS})*100)',
     f_kpi_value_pct),
    ("TOTAL SESSIONS", f'=IF({FILTER}="All",SUM({T_SESS}),SUMIF({T_DIV},{FILTER},{T_SESS}))', f_kpi_value_num),
    ("ACTIVE IOT DEVICES",
     f'=IF({FILTER}="All",SUMIF({I_YM},"{month_list[-1]}",{I_DEV}),SUMIFS({I_DEV},{I_YM},"{month_list[-1]}",{I_DIV},{FILTER}))',
     f_kpi_value_num),
    ("AVG CUSTOMER SATISFACTION (1-5)",
     f'=IF({FILTER}="All",AVERAGE({K_CSAT}),AVERAGEIF({K_DIV},{FILTER},{K_CSAT}))', f_kpi_value_dec),
]
kpi_cols = ["B", "E", "H", "K", "N", "Q"]  # not all used; pick 6 evenly spaced starting cols
kpi_cols = ["B", "F", "J", "B", "F", "J"]
kpi_rows = [6, 6, 6, 10, 10, 10]
for (label, formula, fmt), col, row in zip(kpis, kpi_cols, kpi_rows):
    c0 = col
    c1 = chr(ord(col) + 2)
    dash.merge_range(f"{c0}{row}:{c1}{row}", label, f_kpi_label)
    dash.merge_range(f"{c0}{row+1}:{c1}{row+1}", "", fmt)
    dash.write_formula(f"{c0}{row+1}", formula, fmt)
    dash.set_row(row - 1, 16)
    dash.set_row(row, 30)

dash.write("B14", "", f_note)
dash.merge_range("B14:Q14",
                  "KPI cards respond to the Division filter above. Charts below show full breakdowns across all "
                  "divisions/categories. All values are live formulas over the Data_* tabs — edit the data and "
                  "everything recalculates.", f_note)

# ---------------------------------------------------------------------------
# CHARTS
# ---------------------------------------------------------------------------

def style_chart(c, title):
    c.set_title({"name": title, "name_font": TITLE_FONT})
    c.set_legend({"font": CHART_FONT, "position": "bottom"})
    c.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
    c.set_plotarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
    c.set_x_axis({"num_font": CHART_FONT, "line": {"color": "#CFCFCF"}})
    c.set_y_axis({"num_font": CHART_FONT, "major_gridlines": {"visible": True, "line": {"color": "#EFEFEF"}}})
    c.set_size({"width": 600, "height": 330})


# 1) Combo: Monthly Sessions (col) + Conversion Rate (line, secondary axis)
c1 = wb.add_chart({"type": "column"})
c1.add_series({
    "name": "Sessions", "categories": f"=Calc_Helpers!$A${MT_FIRST}:$A${MT_LAST}",
    "values": f"=Calc_Helpers!$D${MT_FIRST}:$D${MT_LAST}", "fill": {"color": GREEN}, "gap": 30,
})
c1_line = wb.add_chart({"type": "line"})
c1_line.add_series({
    "name": "Conversion Rate %", "categories": f"=Calc_Helpers!$A${MT_FIRST}:$A${MT_LAST}",
    "values": f"=Calc_Helpers!$F${MT_FIRST}:$F${MT_LAST}", "y2_axis": True,
    "line": {"color": TERRA, "width": 2.5}, "marker": {"type": "circle", "size": 6, "fill": {"color": TERRA}},
})
c1.combine(c1_line)
c1.set_y2_axis({"num_font": CHART_FONT, "name": "Conversion Rate (%)", "name_font": CHART_FONT})
style_chart(c1, "Monthly Sessions vs. Conversion Rate (UX Overhaul Impact)")
dash.insert_chart("B16", c1, {"x_offset": 2, "y_offset": 2})

# 2) Line: Monthly Revenue trend
c2 = wb.add_chart({"type": "line"})
c2.add_series({
    "name": "Revenue (BDT)", "categories": f"=Calc_Helpers!$A${MT_FIRST}:$A${MT_LAST}",
    "values": f"=Calc_Helpers!$B${MT_FIRST}:$B${MT_LAST}", "line": {"color": GOLD, "width": 2.5},
    "marker": {"type": "circle", "size": 6, "fill": {"color": GOLD}},
})
style_chart(c2, "Monthly Revenue Trend (BDT)")
dash.insert_chart("J16", c2, {"x_offset": 2, "y_offset": 2})

# 3) Bar: Revenue by Division
c3 = wb.add_chart({"type": "bar"})
c3.add_series({
    "name": "Revenue (BDT)", "categories": f"=Calc_Helpers!$A${DS_FIRST}:$A${DS_LAST}",
    "values": f"=Calc_Helpers!$B${DS_FIRST}:$B${DS_LAST}", "fill": {"color": GREEN},
    "data_labels": {"value": True, "num_format": "#,##0", "font": CHART_FONT},
})
style_chart(c3, "Total Revenue by Division (BDT)")
c3.set_legend({"none": True})
dash.insert_chart("B34", c3, {"x_offset": 2, "y_offset": 2})

# 4) Doughnut: Device model mix (units)
c4 = wb.add_chart({"type": "doughnut"})
c4.add_series({
    "name": "Units Sold", "categories": f"=Calc_Helpers!$A${PS_FIRST}:$A${PS_LAST}",
    "values": f"=Calc_Helpers!$B${PS_FIRST}:$B${PS_LAST}",
    "points": [{"fill": {"color": GREEN}}, {"fill": {"color": GOLD}}, {"fill": {"color": TERRA}}],
    "data_labels": {"percentage": True, "font": {"name": "Carlito", "size": 9, "color": "#FFFFFF"}},
})
c4.set_hole_size(55)
style_chart(c4, "Devices Sold by Model")
dash.insert_chart("J34", c4, {"x_offset": 2, "y_offset": 2})

# 5) Bar combo: Support tickets by category (count) + Avg resolution time (line)
c5 = wb.add_chart({"type": "column"})
c5.add_series({
    "name": "Tickets", "categories": f"=Calc_Helpers!$A${TS_FIRST}:$A${TS_LAST}",
    "values": f"=Calc_Helpers!$B${TS_FIRST}:$B${TS_LAST}", "fill": {"color": "#3E6B8C"}, "gap": 40,
})
c5_line = wb.add_chart({"type": "line"})
c5_line.add_series({
    "name": "Avg Resolution (hrs)", "categories": f"=Calc_Helpers!$A${TS_FIRST}:$A${TS_LAST}",
    "values": f"=Calc_Helpers!$C${TS_FIRST}:$C${TS_LAST}", "y2_axis": True,
    "line": {"color": TERRA, "width": 2.5}, "marker": {"type": "circle", "size": 6, "fill": {"color": TERRA}},
})
c5.combine(c5_line)
c5.set_y2_axis({"num_font": CHART_FONT, "name": "Hours", "name_font": CHART_FONT})
style_chart(c5, "Support Tickets & Avg. Resolution Time by Category")
dash.insert_chart("B52", c5, {"x_offset": 2, "y_offset": 2})

# 6) Bar: IoT alert rate by division
c6 = wb.add_chart({"type": "bar"})
c6.add_series({
    "name": "Alert Rate %", "categories": f"=Calc_Helpers!$A${DS_FIRST}:$A${DS_LAST}",
    "values": f"=Calc_Helpers!$H${DS_FIRST}:$H${DS_LAST}", "fill": {"color": TERRA},
    "data_labels": {"value": True, "num_format": "0.00", "font": CHART_FONT},
})
c6.set_legend({"none": True})
style_chart(c6, "IoT Sensor Alert Rate (%) by Division")
dash.insert_chart("J52", c6, {"x_offset": 2, "y_offset": 2})

# 7) Bar: Top 10 districts by revenue
c7 = wb.add_chart({"type": "bar"})
c7.add_series({
    "name": "Revenue (BDT)", "categories": f"=Calc_Helpers!$A${TD_FIRST}:$A${TD_LAST}",
    "values": f"=Calc_Helpers!$C${TD_FIRST}:$C${TD_LAST}", "fill": {"color": "#7AAE6B"},
})
c7.set_legend({"none": True})
c7.set_y_axis({"reverse": True, "num_font": CHART_FONT})
style_chart(c7, "Top 10 Districts by Revenue (BDT)")
dash.insert_chart("B70", c7, {"x_offset": 2, "y_offset": 2})

# 8) Combo: Channel sessions + conversion rate
c8 = wb.add_chart({"type": "column"})
c8.add_series({
    "name": "Sessions", "categories": f"=Calc_Helpers!$A${CS_FIRST}:$A${CS_LAST}",
    "values": f"=Calc_Helpers!$B${CS_FIRST}:$B${CS_LAST}", "fill": {"color": "#8C5B3E"}, "gap": 40,
})
c8_line = wb.add_chart({"type": "line"})
c8_line.add_series({
    "name": "Conversion Rate %", "categories": f"=Calc_Helpers!$A${CS_FIRST}:$A${CS_LAST}",
    "values": f"=Calc_Helpers!$D${CS_FIRST}:$D${CS_LAST}", "y2_axis": True,
    "line": {"color": GREEN, "width": 2.5}, "marker": {"type": "circle", "size": 6, "fill": {"color": GREEN}},
})
c8.combine(c8_line)
c8.set_y2_axis({"num_font": CHART_FONT, "name": "Conversion Rate (%)", "name_font": CHART_FONT})
style_chart(c8, "Acquisition Channel: Sessions vs. Conversion Rate")
dash.insert_chart("J70", c8, {"x_offset": 2, "y_offset": 2})

dash.freeze_panes(2, 0)
dash.activate()

wb.close()
print("Stage 3 (final) OK:", OUT_PATH)
