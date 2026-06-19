/* =============================================================================
   FairFarm BD — Analytical SQL Queries
   Md. Abul Bashar Nirob | Data & Business Analyst
   -----------------------------------------------------------------------------
   Dialect: SQLite (ANSI-SQL compatible). Runs as-is against data/fairfarm_bd.db,
   and ports to Snowflake/SQL Server/Postgres with minor date-function tweaks).
   Run: sqlite3 data/fairfarm_bd.db < sql/fairfarm_analysis_queries.sql
   ============================================================================= */


-- -----------------------------------------------------------------------------
-- 1. UX OVERHAUL IMPACT — conversion rate & sessions, before vs after launch
--    (validates resume claim: +25% conversion rate, +30% visitor traffic)
-- -----------------------------------------------------------------------------
SELECT
    d.period_flag,
    COUNT(DISTINCT t.date)                                   AS days_in_period,
    SUM(t.sessions)                                          AS total_sessions,
    ROUND(SUM(t.sessions) * 1.0 / COUNT(DISTINCT t.date), 1) AS avg_daily_sessions,
    SUM(t.conversions)                                       AS total_conversions,
    ROUND(SUM(t.conversions) * 100.0 / SUM(t.sessions), 2)   AS conversion_rate_pct
FROM fact_web_traffic t
JOIN dim_date d ON d.date = t.date
GROUP BY d.period_flag;


-- -----------------------------------------------------------------------------
-- 2. MONTHLY REVENUE TREND & MONTH-OVER-MONTH GROWTH
-- -----------------------------------------------------------------------------
WITH monthly_revenue AS (
    SELECT
        strftime('%Y-%m', date) AS year_month,
        SUM(total_amount_bdt)   AS revenue_bdt,
        COUNT(*)                AS orders
    FROM fact_sales
    GROUP BY 1
)
SELECT
    year_month,
    revenue_bdt,
    orders,
    ROUND(
        (revenue_bdt - LAG(revenue_bdt) OVER (ORDER BY year_month)) * 100.0
        / LAG(revenue_bdt) OVER (ORDER BY year_month), 1
    ) AS mom_growth_pct
FROM monthly_revenue
ORDER BY year_month;


-- -----------------------------------------------------------------------------
-- 3. REGIONAL SALES PERFORMANCE — revenue, orders & avg order value by division
-- -----------------------------------------------------------------------------
SELECT
    division_name,
    COUNT(*)                                   AS total_orders,
    SUM(quantity)                               AS units_sold,
    SUM(total_amount_bdt)                       AS total_revenue_bdt,
    ROUND(AVG(total_amount_bdt), 0)             AS avg_order_value_bdt,
    ROUND(SUM(total_amount_bdt) * 100.0 / (SELECT SUM(total_amount_bdt) FROM fact_sales), 1) AS pct_of_total_revenue
FROM fact_sales
GROUP BY division_name
ORDER BY total_revenue_bdt DESC;


-- -----------------------------------------------------------------------------
-- 4. PRODUCT MIX — devices sold & revenue contribution by model
-- -----------------------------------------------------------------------------
SELECT
    device_model,
    SUM(quantity)                                                              AS units_sold,
    SUM(total_amount_bdt)                                                      AS revenue_bdt,
    ROUND(SUM(total_amount_bdt) * 100.0 / (SELECT SUM(total_amount_bdt) FROM fact_sales), 1) AS pct_of_revenue
FROM fact_sales
GROUP BY device_model
ORDER BY revenue_bdt DESC;


-- -----------------------------------------------------------------------------
-- 5. TOP 10 DISTRICTS BY REVENUE (for dealer / inventory planning)
-- -----------------------------------------------------------------------------
SELECT
    district_name,
    division_name,
    COUNT(*)               AS total_orders,
    SUM(total_amount_bdt)  AS total_revenue_bdt
FROM fact_sales
GROUP BY district_name, division_name
ORDER BY total_revenue_bdt DESC
LIMIT 10;


-- -----------------------------------------------------------------------------
-- 6. CUSTOMER SEGMENTATION — farm size segment vs spend & device adoption
-- -----------------------------------------------------------------------------
SELECT
    c.customer_segment,
    COUNT(DISTINCT c.customer_id)                AS customers,
    COUNT(s.order_id)                            AS orders,
    ROUND(SUM(s.total_amount_bdt), 0)            AS total_revenue_bdt,
    ROUND(SUM(s.total_amount_bdt) * 1.0 / COUNT(DISTINCT c.customer_id), 0) AS revenue_per_customer_bdt
FROM dim_customer c
LEFT JOIN fact_sales s ON s.customer_id = c.customer_id
GROUP BY c.customer_segment
ORDER BY revenue_per_customer_bdt DESC;


-- -----------------------------------------------------------------------------
-- 7. IOT DEVICE HEALTH — alert rate & breakdown by alert type, by division
-- -----------------------------------------------------------------------------
SELECT
    division_name,
    COUNT(*)                                                  AS total_readings,
    SUM(CASE WHEN alert_flag = 1 THEN 1 ELSE 0 END)            AS alert_readings,
    ROUND(SUM(CASE WHEN alert_flag = 1 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2) AS alert_rate_pct,
    ROUND(AVG(soil_moisture_pct), 1)                           AS avg_soil_moisture_pct,
    ROUND(AVG(soil_ph), 2)                                     AS avg_soil_ph
FROM fact_iot_readings
GROUP BY division_name
ORDER BY alert_rate_pct DESC;


-- -----------------------------------------------------------------------------
-- 8. SUPPORT SLA — avg resolution time & CSAT by category, pre vs post process fix
-- -----------------------------------------------------------------------------
SELECT
    category,
    CASE WHEN date >= '2025-12-01' THEN 'Post-Process Fix' ELSE 'Pre-Process Fix' END AS period,
    COUNT(*)                          AS tickets,
    ROUND(AVG(resolution_time_hrs), 1) AS avg_resolution_hrs,
    ROUND(AVG(csat_score), 2)          AS avg_csat
FROM fact_support_tickets
GROUP BY category, period
ORDER BY category, period;


-- -----------------------------------------------------------------------------
-- 9. COHORT-STYLE VIEW — monthly new customer signups vs that month's revenue
-- -----------------------------------------------------------------------------
SELECT
    strftime('%Y-%m', c.signup_date)         AS signup_month,
    COUNT(DISTINCT c.customer_id)            AS new_customers,
    ROUND(COALESCE(SUM(s.total_amount_bdt), 0), 0) AS revenue_from_orders_same_month
FROM dim_customer c
LEFT JOIN fact_sales s
    ON s.customer_id = c.customer_id
   AND strftime('%Y-%m', s.date) = strftime('%Y-%m', c.signup_date)
GROUP BY signup_month
ORDER BY signup_month;


-- -----------------------------------------------------------------------------
-- 10. ACQUISITION CHANNEL EFFICIENCY — which channel brings highest-value customers
-- -----------------------------------------------------------------------------
SELECT
    c.acquisition_channel,
    COUNT(DISTINCT c.customer_id)                                 AS customers,
    ROUND(SUM(s.total_amount_bdt), 0)                             AS total_revenue_bdt,
    ROUND(SUM(s.total_amount_bdt) * 1.0 / COUNT(DISTINCT c.customer_id), 0) AS revenue_per_customer_bdt
FROM dim_customer c
LEFT JOIN fact_sales s ON s.customer_id = c.customer_id
GROUP BY c.acquisition_channel
ORDER BY revenue_per_customer_bdt DESC;
