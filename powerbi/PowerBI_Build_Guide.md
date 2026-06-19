# Power BI Build Guide — FairFarm BD Matir Doctor Dashboard

> **Why isn't there a ready-made `.pbix` file in this repo?**
> A `.pbix` is a Windows-only, proprietary binary (it embeds a compressed
> Vertipaq data model) that can only be authored inside **Power BI Desktop**.
> This project was built in a Linux environment that can produce the dataset,
> SQL, Excel dashboard and every Power BI building block — but not the
> `.pbix` binary itself. The good news: with the kit in this folder, you can
> assemble the real, working `.pbix` in well under an hour. Follow the steps
> below in order.

**What's in this folder:**
| File | Purpose |
|---|---|
| `PowerQuery_M_Scripts.txt` | Paste-in M code to import & type every table |
| `DAX_Measures.txt` | All 25+ DAX measures, ready to paste |
| `FairFarm_BD_Theme.json` | Custom color theme — import directly in Power BI |
| `PowerBI_Build_Guide.md` | This file |

---

## Step 1 — Get the data onto your machine

Unzip the project and note the full path to `data/csv/` — you'll need it in Step 2.

## Step 2 — Import the tables

1. Open **Power BI Desktop** → **Get Data** → **Text/CSV**, OR use the faster
   route below with a parameter so all queries share one folder path.
2. **Home → Manage Parameters → New Parameter**
   - Name: `FolderPath`, Type: `Text`
   - Current Value: the full path to `data/csv/`, ending in a slash
     (e.g. `C:\Users\you\fairfarm-bd-analytics\data\csv\`)
3. For each table — `dim_date`, `dim_division`, `dim_district`, `dim_product`,
   `dim_customer`, `fact_sales`, `fact_web_traffic`, `fact_iot_readings`,
   `fact_support_tickets`, `stakeholder_gap_analysis` — create a **Blank Query**
   (Home → Get Data → Blank Query), open the **Advanced Editor**, and paste the
   matching block from `PowerQuery_M_Scripts.txt`. Rename each query to match
   the table name exactly (Power BI names new queries "Query1" by default).
4. **Home → Close & Apply**.

## Step 3 — Build the data model (star schema)

Go to **Model view** and create these relationships (all 1-to-many, single
direction, from dimension to fact):

```
dim_date[date]            1 ──── *  fact_sales[date]
dim_date[date]            1 ──── *  fact_web_traffic[date]
dim_date[date]            1 ──── *  fact_iot_readings[date]
dim_date[date]            1 ──── *  fact_support_tickets[date]

dim_division[division_id] 1 ──── *  fact_sales[division_id]
dim_division[division_id] 1 ──── *  fact_web_traffic[division_id]
dim_division[division_id] 1 ──── *  fact_iot_readings[division_id]
dim_division[division_id] 1 ──── *  fact_support_tickets[division_id]

dim_product[product_id]   1 ──── *  fact_sales[product_id]
dim_customer[customer_id] 1 ──── *  fact_sales[customer_id]
dim_customer[customer_id] 1 ──── *  fact_support_tickets[customer_id]
dim_customer[customer_id] 1 ──── *  fact_iot_readings[customer_id]
```

Mark `dim_date` as a **Date Table**: select the table → Table tools →
**Mark as Date Table** → pick the `date` column.

## Step 4 — Add the DAX measures

Create a new blank table to hold measures so they don't clutter your fact
tables: **Home → Enter Data**, name it `_Measures`, click Load with no rows.
Then, for every block in `DAX_Measures.txt`, go to **Modeling → New Measure**,
paste the formula, and set the home table to `_Measures`.
Set number formats where helpful (Currency/BDT, Percentage, Decimal).

## Step 5 — Apply the custom theme

**View → Themes → Browse for themes…** → select `FairFarm_BD_Theme.json`.
This applies the project's signature forest-green / gold / terracotta palette
used throughout this repo (Excel dashboard, charts, README).

## Step 6 — Build the report pages

Recommended 3-page layout (mirrors `images/dashboard_preview.png`):

### Page 1 — Executive Overview
- **Card visuals** (top row): `[Total Revenue (BDT)]`, `[Devices Sold (Units)]`,
  `[Conversion Rate %]`, `[Total Sessions]`, `[Active IoT Devices]`,
  `[Avg CSAT Score]`
- **Combo chart** (column + line, secondary axis): X = `dim_date[year_month]`,
  Columns = `[Total Sessions]`, Line = `[Conversion Rate %]`. Add a vertical
  reference line / shape at Dec 2025 labelled "UX overhaul shipped" to call
  out the before/after story.
- **Clustered bar**: Revenue by `dim_division[division_name]`
- **Donut chart**: Devices Sold by `dim_product[device_model]`
- **Slicer**: `dim_division[division_name]` (sync across all pages)

### Page 2 — Regional & Product Deep-Dive
- **Map / filled map visual**: `dim_division[latitude]` / `[longitude]`, size =
  `[Total Revenue (BDT)]` (Bangladesh divisions)
- **Table**: Top 10 districts by `[Total Revenue (BDT)]`
- **Bar chart**: Revenue per Customer by `dim_customer[customer_segment]`
- **Bar chart**: Revenue per Customer by `dim_customer[acquisition_channel]`

### Page 3 — IoT Device Health & Support Operations
- **Line chart**: `[Avg Soil Moisture %]` and `[Avg Soil pH]` by month
- **Bar chart**: `[IoT Alert Rate %]` by division
- **Bar + line combo**: Ticket count + `[Avg Resolution Time (hrs)]` by
  `fact_support_tickets[category]`
- **Card**: `[Resolution Time Improvement %]` to show the process-fix impact
- **Table visual**: the `stakeholder_gap_analysis` table — a nice touch that
  shows the requirements/BRD work behind the numbers

## Step 7 — Polish

- Page background: white/cream (`#FBF8F1`) to match the theme
- Title text box per page using the brand header copy from
  `images/dashboard_preview.png`
- Format all currency fields as `"BDT "#,##0`
- Add tooltips/bookmarks if you want a guided walkthrough

## Step 8 — Save and export

**File → Save As → FairFarm_BD_Dashboard.pbix**. You're done — and the file
you just produced is a genuine, fully-interactive Power BI report you can
publish to the Power BI Service or share as-is.

---

### Quick sanity checks once built
- Total Revenue card should read **≈ BDT 29.7M**
- Conversion Rate Lift (Post-UX vs Pre-UX) should land **≈ +25%**
- Visitor Traffic Lift should land **≈ +30%**

These match the resume's documented results and the numbers already verified
in `docs/eda_kpi_summary.json` and the SQL queries in `sql/`.
