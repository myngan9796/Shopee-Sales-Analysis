/* =============================================================================
   SHOPEE 2025 SALES PERFORMANCE — FULL SQL ANALYSIS
   -----------------------------------------------------------------------------
   Purpose   : Complete, standalone SQL companion to the "Shopee 2025 Sales
               Performance" data analyst case study. Every metric quoted in the
               Word/PDF report, the HTML portfolio page, and the interactive
               dashboard is reproduced here as an auditable query.
   Dialect   : PostgreSQL (uses EXTRACT, TO_CHAR, CEIL, window functions, CTEs).
               Tested for logical correctness against the source dataset using
               SQLite 3.45 as a validation engine; syntax below is native
               PostgreSQL. Minor notes are left where a dialect swap would be
               needed (e.g. DATE_TRUNC vs EXTRACT, ILIKE vs LIKE).
   Source    : "2025_Sale_Performance.xlsx" — Sales Overview sheet (365 daily
               rows, 2025-01-01 to 2025-12-31, one row per calendar day).
   Author    : Data Analyst portfolio case study
   ============================================================================= */


/* =============================================================================
   SECTION 0 — SCHEMA & TABLE SETUP
   ============================================================================= */

-- 0.1  Base table: one row per calendar day, raw counters only (no derived
--      rates stored — everything downstream is computed in SQL so the numbers
--      never drift out of sync with the source counters).
DROP TABLE IF EXISTS shopee_sales_daily;

CREATE TABLE shopee_sales_daily (
    sales_date          DATE            NOT NULL PRIMARY KEY,
    visitors             BIGINT         NOT NULL,   -- unique visits that day
    buyers_placed         BIGINT        NOT NULL,   -- distinct buyers, placed orders
    units_placed           BIGINT       NOT NULL,   -- units, placed orders
    orders_placed           BIGINT      NOT NULL,   -- order count, placed orders
    gmv_placed_vnd            NUMERIC(18,2) NOT NULL, -- GMV (VND), placed orders
    buyers_confirmed       BIGINT        NOT NULL,   -- distinct buyers, confirmed orders
    units_confirmed         BIGINT       NOT NULL,   -- units, confirmed orders
    orders_confirmed         BIGINT      NOT NULL,   -- order count, confirmed orders
    gmv_confirmed_vnd          NUMERIC(18,2) NOT NULL -- GMV (VND), confirmed orders
);

-- 0.2  Load data (adjust path/COPY options to your environment):
-- COPY shopee_sales_daily (sales_date, visitors, buyers_placed, units_placed,
--   orders_placed, gmv_placed_vnd, buyers_confirmed, units_confirmed,
--   orders_confirmed, gmv_confirmed_vnd)
-- FROM '/path/to/daily.csv'
-- WITH (FORMAT csv, HEADER true);

CREATE INDEX IF NOT EXISTS idx_ssd_date ON shopee_sales_daily (sales_date);

-- 0.3  Reusable "enriched" view — adds calendar attributes and the campaign-day
--      flag that every downstream section joins against. "Confirmed orders"
--      is the business's primary GMV metric (net of buyer/seller cancellations
--      and returns caught before confirmation), so it is used as the default
--      basis throughout unless a query explicitly says "placed".
--
--      Campaign-day rule: Shopee's 2025 flash-sale calendar is the twelve
--      "double-date" days (1.1, 2.2, 3.3, ... 12.12) — i.e. the day-of-month
--      equals the month number. This matches the 12 rows in the case study's
--      campaign-uplift table exactly.
CREATE OR REPLACE VIEW v_sales_daily AS
SELECT
    d.*,
    EXTRACT(MONTH FROM d.sales_date)::int                              AS sales_month,
    CEIL(EXTRACT(MONTH FROM d.sales_date) / 3.0)::int                  AS sales_quarter,
    CASE WHEN EXTRACT(MONTH FROM d.sales_date) <= 6 THEN 'H1' ELSE 'H2' END AS half_year,
    TRIM(TO_CHAR(d.sales_date, 'Day'))                                  AS day_of_week_name,
    EXTRACT(DOW FROM d.sales_date)::int                                 AS day_of_week_num,   -- 0=Sun ... 6=Sat
    CASE WHEN EXTRACT(DAY FROM d.sales_date) = EXTRACT(MONTH FROM d.sales_date)
         THEN 1 ELSE 0 END                                              AS is_campaign_day
FROM shopee_sales_daily d;


/* =============================================================================
   SECTION 1 — DATA QUALITY CHECKS
   ============================================================================= */

-- 1.1  Row count & date-range sanity check — expect 365 rows, Jan 1 to Dec 31.
SELECT
    COUNT(*)                AS row_count,
    MIN(sales_date)          AS first_date,
    MAX(sales_date)          AS last_date,
    MAX(sales_date) - MIN(sales_date) + 1 AS expected_days
FROM shopee_sales_daily;

-- 1.2  Duplicate dates — expect 0 rows returned.
SELECT sales_date, COUNT(*) AS n
FROM shopee_sales_daily
GROUP BY sales_date
HAVING COUNT(*) > 1;

-- 1.3  Missing calendar days — left-join against a generated date series to
--      surface any gap in the 365-day sequence. Expect 0 rows.
SELECT gs.d AS missing_date
FROM generate_series(DATE '2025-01-01', DATE '2025-12-31', INTERVAL '1 day') AS gs(d)
LEFT JOIN shopee_sales_daily s ON s.sales_date = gs.d::date
WHERE s.sales_date IS NULL;

-- 1.4  Null / negative value checks on every numeric column.
SELECT
    SUM(CASE WHEN visitors            IS NULL OR visitors            < 0 THEN 1 ELSE 0 END) AS bad_visitors,
    SUM(CASE WHEN buyers_placed       IS NULL OR buyers_placed       < 0 THEN 1 ELSE 0 END) AS bad_buyers_placed,
    SUM(CASE WHEN units_placed        IS NULL OR units_placed        < 0 THEN 1 ELSE 0 END) AS bad_units_placed,
    SUM(CASE WHEN orders_placed       IS NULL OR orders_placed       < 0 THEN 1 ELSE 0 END) AS bad_orders_placed,
    SUM(CASE WHEN gmv_placed_vnd      IS NULL OR gmv_placed_vnd      < 0 THEN 1 ELSE 0 END) AS bad_gmv_placed,
    SUM(CASE WHEN buyers_confirmed    IS NULL OR buyers_confirmed    < 0 THEN 1 ELSE 0 END) AS bad_buyers_confirmed,
    SUM(CASE WHEN units_confirmed     IS NULL OR units_confirmed     < 0 THEN 1 ELSE 0 END) AS bad_units_confirmed,
    SUM(CASE WHEN orders_confirmed    IS NULL OR orders_confirmed    < 0 THEN 1 ELSE 0 END) AS bad_orders_confirmed,
    SUM(CASE WHEN gmv_confirmed_vnd   IS NULL OR gmv_confirmed_vnd   < 0 THEN 1 ELSE 0 END) AS bad_gmv_confirmed
FROM shopee_sales_daily;

-- 1.5  Logical-consistency check: confirmed counters should generally not
--      exceed placed counters *within the same day*. A small number of
--      exceptions is EXPECTED and not an error — Shopee confirms orders on a
--      rolling basis, so a day's "confirmed" total can include orders that
--      were placed on a previous day and only cleared confirmation today.
--      This query surfaces those days for awareness, not for "fixing".
SELECT
    sales_date,
    buyers_placed, buyers_confirmed,
    units_placed,  units_confirmed,
    orders_placed, orders_confirmed
FROM shopee_sales_daily
WHERE buyers_confirmed > buyers_placed
   OR units_confirmed  > units_placed
   OR orders_confirmed > orders_placed
ORDER BY sales_date;

-- 1.6  Outlier scan — flag days whose confirmed GMV is more than 3 standard
--      deviations from the annual mean (campaign days are expected to surface
--      here; this is a sense-check, not a cleaning step).
WITH stats AS (
    SELECT AVG(gmv_confirmed_vnd) AS mean_gmv, STDDEV_SAMP(gmv_confirmed_vnd) AS sd_gmv
    FROM shopee_sales_daily
)
SELECT s.sales_date, s.gmv_confirmed_vnd,
       ROUND((s.gmv_confirmed_vnd - st.mean_gmv) / st.sd_gmv, 2) AS z_score
FROM shopee_sales_daily s CROSS JOIN stats st
WHERE ABS(s.gmv_confirmed_vnd - st.mean_gmv) > 3 * st.sd_gmv
ORDER BY s.sales_date;


/* =============================================================================
   SECTION 2 — ANNUAL KPI OVERVIEW
   ============================================================================= */

-- 2.1  Headline annual KPIs (matches Executive Summary / Table 5.1 of the report).
SELECT
    SUM(visitors)                                                  AS total_visitors,
    SUM(buyers_confirmed)                                          AS total_confirmed_buyers,
    SUM(gmv_confirmed_vnd)                                         AS total_confirmed_gmv_vnd,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(visitors), 2)        AS overall_conversion_rate_pct,
    ROUND(SUM(gmv_confirmed_vnd) / SUM(buyers_confirmed), 0)       AS average_order_value_vnd,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(buyers_placed), 2)   AS overall_confirmation_rate_pct
FROM shopee_sales_daily;

-- 2.2  Placed vs. confirmed side-by-side (basis comparison used throughout
--      the dashboard's "Placed vs Confirmed" filter).
SELECT
    'Placed'    AS basis, SUM(visitors) AS visitors, SUM(buyers_placed)    AS buyers, SUM(units_placed)    AS units, SUM(orders_placed)    AS orders, SUM(gmv_placed_vnd)    AS gmv_vnd
FROM shopee_sales_daily
UNION ALL
SELECT
    'Confirmed' AS basis, SUM(visitors) AS visitors, SUM(buyers_confirmed) AS buyers, SUM(units_confirmed) AS units, SUM(orders_confirmed) AS orders, SUM(gmv_confirmed_vnd) AS gmv_vnd
FROM shopee_sales_daily;


/* =============================================================================
   SECTION 3 — MONTHLY TREND & MONTH-OVER-MONTH (MoM) GROWTH
   ============================================================================= */

WITH monthly AS (
    SELECT
        sales_month,
        SUM(visitors)          AS visitors,
        SUM(buyers_confirmed)  AS buyers_confirmed,
        SUM(gmv_confirmed_vnd) AS gmv_confirmed_vnd
    FROM v_sales_daily
    GROUP BY sales_month
)
SELECT
    sales_month,
    visitors,
    buyers_confirmed,
    gmv_confirmed_vnd,
    ROUND(100.0 * buyers_confirmed / visitors, 2)                                    AS conversion_rate_pct,
    ROUND(gmv_confirmed_vnd / buyers_confirmed, 0)                                    AS aov_vnd,
    ROUND(
        100.0 * (gmv_confirmed_vnd - LAG(gmv_confirmed_vnd) OVER (ORDER BY sales_month))
        / NULLIF(LAG(gmv_confirmed_vnd) OVER (ORDER BY sales_month), 0), 2
    )                                                                                  AS mom_growth_pct
FROM monthly
ORDER BY sales_month;


/* =============================================================================
   SECTION 4 — QUARTERLY TREND & QUARTER-OVER-QUARTER (QoQ) GROWTH
   ============================================================================= */

WITH quarterly AS (
    SELECT
        sales_quarter,
        SUM(visitors)          AS visitors,
        SUM(buyers_confirmed)  AS buyers_confirmed,
        SUM(gmv_confirmed_vnd) AS gmv_confirmed_vnd
    FROM v_sales_daily
    GROUP BY sales_quarter
)
SELECT
    sales_quarter,
    visitors,
    buyers_confirmed,
    gmv_confirmed_vnd,
    ROUND(100.0 * buyers_confirmed / visitors, 2)  AS conversion_rate_pct,
    ROUND(gmv_confirmed_vnd / buyers_confirmed, 0) AS aov_vnd,
    ROUND(
        100.0 * (gmv_confirmed_vnd - LAG(gmv_confirmed_vnd) OVER (ORDER BY sales_quarter))
        / NULLIF(LAG(gmv_confirmed_vnd) OVER (ORDER BY sales_quarter), 0), 2
    )                                                AS qoq_growth_pct
FROM quarterly
ORDER BY sales_quarter;


/* =============================================================================
   SECTION 5 — H1 vs H2 COMPARISON
   ============================================================================= */

WITH halves AS (
    SELECT
        half_year,
        SUM(visitors)          AS visitors,
        SUM(buyers_confirmed)  AS buyers_confirmed,
        SUM(gmv_confirmed_vnd) AS gmv_confirmed_vnd
    FROM v_sales_daily
    GROUP BY half_year
)
SELECT
    half_year,
    visitors,
    buyers_confirmed,
    gmv_confirmed_vnd,
    ROUND(100.0 * buyers_confirmed / visitors, 2)  AS conversion_rate_pct,
    ROUND(gmv_confirmed_vnd / buyers_confirmed, 0) AS aov_vnd,
    ROUND(
        100.0 * (gmv_confirmed_vnd - LAG(gmv_confirmed_vnd) OVER (ORDER BY half_year))
        / NULLIF(LAG(gmv_confirmed_vnd) OVER (ORDER BY half_year), 0), 2
    )                                                AS h2_vs_h1_growth_pct
FROM halves
ORDER BY half_year;


/* =============================================================================
   SECTION 6 — CAMPAIGN DAY (FLASH SALE) ANALYSIS
   ============================================================================= */

-- 6.1  Per-campaign-day detail: GMV, visitors, conversion rate vs. that
--      month's all-day average, and the resulting uplift multiple.
--      (This reproduces campaign_days.csv / Figure 2 exactly.)
WITH month_avg AS (
    SELECT sales_month, AVG(gmv_confirmed_vnd) AS avg_gmv_vnd, AVG(visitors) AS avg_visitors, AVG(100.0*buyers_confirmed/visitors) AS avg_cr_pct
    FROM v_sales_daily
    GROUP BY sales_month
)
SELECT
    v.sales_date,
    v.visitors,
    v.buyers_confirmed,
    v.gmv_confirmed_vnd,
    ROUND(100.0 * v.buyers_confirmed / v.visitors, 2)                    AS conversion_rate_pct,
    ROUND(v.gmv_confirmed_vnd / v.buyers_confirmed, 0)                    AS aov_vnd,
    ROUND(v.gmv_confirmed_vnd / m.avg_gmv_vnd, 2)                         AS gmv_uplift_x,
    ROUND(v.visitors / m.avg_visitors, 2)                                 AS visitor_uplift_x,
    ROUND((100.0*v.buyers_confirmed/v.visitors) / m.avg_cr_pct, 2)        AS cr_uplift_x
FROM v_sales_daily v
JOIN month_avg m ON m.sales_month = v.sales_month
WHERE v.is_campaign_day = 1
ORDER BY v.sales_date;

-- 6.2  Campaign vs. non-campaign contribution to annual GMV — quantifies how
--      much of the year's GMV comes from just 12 flash-sale days.
SELECT
    is_campaign_day,
    COUNT(*)                                                       AS n_days,
    SUM(gmv_confirmed_vnd)                                         AS gmv_confirmed_vnd,
    ROUND(100.0 * SUM(gmv_confirmed_vnd) / SUM(SUM(gmv_confirmed_vnd)) OVER (), 2) AS pct_of_annual_gmv,
    ROUND(AVG(gmv_confirmed_vnd), 0)                               AS avg_daily_gmv_vnd
FROM v_sales_daily
GROUP BY is_campaign_day
ORDER BY is_campaign_day;

-- 6.3  Ranking campaign days by uplift multiple — which flash sale over-
--      performed its own month the most.
WITH month_avg AS (
    SELECT sales_month, AVG(gmv_confirmed_vnd) AS avg_gmv_vnd
    FROM v_sales_daily
    GROUP BY sales_month
)
SELECT
    v.sales_date,
    v.gmv_confirmed_vnd,
    ROUND(v.gmv_confirmed_vnd / m.avg_gmv_vnd, 2) AS gmv_uplift_x,
    RANK() OVER (ORDER BY v.gmv_confirmed_vnd / m.avg_gmv_vnd DESC) AS uplift_rank
FROM v_sales_daily v
JOIN month_avg m ON m.sales_month = v.sales_month
WHERE v.is_campaign_day = 1
ORDER BY uplift_rank;


/* =============================================================================
   SECTION 7 — DAY-OF-WEEK ANALYSIS
   ============================================================================= */

-- 7.1  Raw day-of-week averages INCLUDING campaign days. Note: because 5 of
--      the 12 campaign days in the 2025 calendar (4.4, 6.6, 8.8, 10.10, 12.12)
--      fall on a Friday, their outsized GMV skews Friday's raw average upward
--      — see 7.2 for the organic (non-campaign) pattern used in the report.
SELECT
    day_of_week_num,
    day_of_week_name,
    COUNT(*)                          AS n_days,
    ROUND(AVG(gmv_confirmed_vnd), 0)  AS avg_gmv_vnd,
    ROUND(AVG(visitors), 0)           AS avg_visitors
FROM v_sales_daily
GROUP BY day_of_week_num, day_of_week_name
ORDER BY day_of_week_num;

-- 7.2  Day-of-week averages EXCLUDING campaign days — the "organic" weekly
--      pattern quoted in the report (Thursday highest, Friday lowest).
SELECT
    day_of_week_num,
    day_of_week_name,
    COUNT(*)                          AS n_days,
    ROUND(AVG(gmv_confirmed_vnd), 0)  AS avg_gmv_vnd,
    ROUND(AVG(visitors), 0)           AS avg_visitors,
    RANK() OVER (ORDER BY AVG(gmv_confirmed_vnd) DESC) AS gmv_rank
FROM v_sales_daily
WHERE is_campaign_day = 0
GROUP BY day_of_week_num, day_of_week_name
ORDER BY day_of_week_num;


/* =============================================================================
   SECTION 8 — CONVERSION FUNNEL & CANCELLATION ANALYSIS
   ============================================================================= */

-- 8.1  Annual funnel: Visitors -> Placed buyers -> Confirmed buyers.
SELECT
    SUM(visitors)                                             AS total_visitors,
    SUM(buyers_placed)                                        AS total_buyers_placed,
    SUM(buyers_confirmed)                                     AS total_buyers_confirmed,
    ROUND(100.0 * SUM(buyers_placed)    / SUM(visitors), 2)   AS cr_visit_to_placed_pct,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(visitors), 2)   AS cr_visit_to_confirmed_pct
FROM shopee_sales_daily;

-- 8.2  Cancellation / drop-off rate at every granularity — buyer, order,
--      unit, and GMV level. This is the query behind the "13.1% / 13.5% /
--      6.9% / 13.8%" figures explained earlier in the case study: cancellation
--      looks different depending on whether you count buyers, orders, units,
--      or value, because higher-value orders cancel disproportionately more.
SELECT
    ROUND(100.0 * (SUM(buyers_placed) - SUM(buyers_confirmed)) / SUM(buyers_placed), 2) AS buyer_dropoff_pct,
    ROUND(100.0 * (SUM(orders_placed) - SUM(orders_confirmed)) / SUM(orders_placed), 2) AS order_dropoff_pct,
    ROUND(100.0 * (SUM(units_placed)  - SUM(units_confirmed))  / SUM(units_placed), 2)  AS unit_dropoff_pct,
    ROUND(100.0 * (SUM(gmv_placed_vnd) - SUM(gmv_confirmed_vnd)) / SUM(gmv_placed_vnd), 2) AS gmv_dropoff_pct
FROM shopee_sales_daily;

-- 8.3  Monthly cancellation trend at GMV level — is drop-off improving or
--      worsening over the year?
SELECT
    sales_month,
    ROUND(100.0 * (SUM(gmv_placed_vnd) - SUM(gmv_confirmed_vnd)) / SUM(gmv_placed_vnd), 2) AS gmv_dropoff_pct,
    ROUND(100.0 * (SUM(orders_placed) - SUM(orders_confirmed)) / SUM(orders_placed), 2)    AS order_dropoff_pct
FROM v_sales_daily
GROUP BY sales_month
ORDER BY sales_month;

-- 8.4  Implied average value of a cancelled order vs. a confirmed order —
--      the arithmetic behind "higher-value orders are disproportionately
--      more likely to cancel."
WITH totals AS (
    SELECT
        SUM(orders_placed) - SUM(orders_confirmed)       AS cancelled_orders,
        SUM(gmv_placed_vnd) - SUM(gmv_confirmed_vnd)      AS cancelled_gmv_vnd,
        SUM(orders_confirmed)                              AS confirmed_orders,
        SUM(gmv_confirmed_vnd)                             AS confirmed_gmv_vnd
    FROM shopee_sales_daily
)
SELECT
    ROUND(cancelled_gmv_vnd / NULLIF(cancelled_orders, 0), 0) AS avg_value_per_cancelled_order_vnd,
    ROUND(confirmed_gmv_vnd / NULLIF(confirmed_orders, 0), 0) AS avg_value_per_confirmed_order_vnd
FROM totals;


/* =============================================================================
   SECTION 9 — ORDER CONFIRMATION RATE TREND (Placed -> Confirmed)
   ============================================================================= */

SELECT
    sales_month,
    SUM(buyers_placed)                                          AS buyers_placed,
    SUM(buyers_confirmed)                                       AS buyers_confirmed,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(buyers_placed), 2) AS confirmation_rate_pct,
    ROUND(
        100.0 * SUM(buyers_confirmed) / SUM(buyers_placed)
        - LAG(100.0 * SUM(buyers_confirmed) / SUM(buyers_placed)) OVER (ORDER BY sales_month), 2
    )                                                             AS ppt_change_vs_prior_month
FROM v_sales_daily
GROUP BY sales_month
ORDER BY sales_month;


/* =============================================================================
   SECTION 10 — SALES PER BUYER (AOV) TREND
   ============================================================================= */

SELECT
    sales_month,
    SUM(gmv_confirmed_vnd)                            AS gmv_confirmed_vnd,
    SUM(buyers_confirmed)                             AS buyers_confirmed,
    ROUND(SUM(gmv_confirmed_vnd) / SUM(buyers_confirmed), 0) AS aov_vnd,
    ROUND(
        100.0 * (SUM(gmv_confirmed_vnd) / SUM(buyers_confirmed)
                 - LAG(SUM(gmv_confirmed_vnd) / SUM(buyers_confirmed)) OVER (ORDER BY sales_month))
        / NULLIF(LAG(SUM(gmv_confirmed_vnd) / SUM(buyers_confirmed)) OVER (ORDER BY sales_month), 0), 2
    )                                                   AS mom_aov_growth_pct
FROM v_sales_daily
GROUP BY sales_month
ORDER BY sales_month;


/* =============================================================================
   SECTION 11 — CONVERSION RATE TREND (Visit -> Confirmed Buyer)
   ============================================================================= */

SELECT
    sales_month,
    SUM(visitors)                                            AS visitors,
    SUM(buyers_confirmed)                                    AS buyers_confirmed,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(visitors), 2)  AS conversion_rate_pct,
    ROUND(
        100.0 * SUM(buyers_confirmed) / SUM(visitors)
        - LAG(100.0 * SUM(buyers_confirmed) / SUM(visitors)) OVER (ORDER BY sales_month), 2
    )                                                          AS ppt_change_vs_prior_month
FROM v_sales_daily
GROUP BY sales_month
ORDER BY sales_month;


/* =============================================================================
   SECTION 12 — TOP-N / RANKING QUERIES
   ============================================================================= */

-- 12.1  Top 10 single days by confirmed GMV (expect all 12 campaign days to
--       dominate the top of this list).
SELECT
    sales_date,
    gmv_confirmed_vnd,
    RANK() OVER (ORDER BY gmv_confirmed_vnd DESC) AS gmv_rank
FROM shopee_sales_daily
ORDER BY gmv_confirmed_vnd DESC
LIMIT 10;

-- 12.2  Bottom 10 single days by confirmed GMV (weakest trading days).
SELECT sales_date, gmv_confirmed_vnd
FROM shopee_sales_daily
ORDER BY gmv_confirmed_vnd ASC
LIMIT 10;

-- 12.3  Best non-campaign day each month (organic performance ceiling,
--       excluding flash-sale spikes).
WITH ranked AS (
    SELECT
        sales_date, sales_month, gmv_confirmed_vnd,
        ROW_NUMBER() OVER (PARTITION BY sales_month ORDER BY gmv_confirmed_vnd DESC) AS rn
    FROM v_sales_daily
    WHERE is_campaign_day = 0
)
SELECT sales_month, sales_date, gmv_confirmed_vnd
FROM ranked
WHERE rn = 1
ORDER BY sales_month;

-- 12.4  Top 5 months by conversion rate.
SELECT
    sales_month,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(visitors), 2) AS conversion_rate_pct
FROM v_sales_daily
GROUP BY sales_month
ORDER BY conversion_rate_pct DESC
LIMIT 5;


/* =============================================================================
   SECTION 13 — ROLLING / MOVING AVERAGE ANALYSIS
   ============================================================================= */

-- 13.1  7-day trailing rolling average of confirmed GMV — smooths daily
--       noise and campaign spikes to show the underlying trend line.
SELECT
    sales_date,
    gmv_confirmed_vnd,
    ROUND(
        AVG(gmv_confirmed_vnd) OVER (ORDER BY sales_date ROWS BETWEEN 6 PRECEDING AND CURRENT ROW), 0
    ) AS rolling_7d_avg_gmv_vnd
FROM shopee_sales_daily
ORDER BY sales_date;

-- 13.2  30-day trailing rolling average of visitors — longer-window traffic
--       trend, less sensitive to any single campaign day.
SELECT
    sales_date,
    visitors,
    ROUND(
        AVG(visitors) OVER (ORDER BY sales_date ROWS BETWEEN 29 PRECEDING AND CURRENT ROW), 0
    ) AS rolling_30d_avg_visitors
FROM shopee_sales_daily
ORDER BY sales_date;


/* =============================================================================
   SECTION 14 — CUMULATIVE / YEAR-TO-DATE (YTD) ANALYSIS
   ============================================================================= */

-- 14.1  Running (YTD) total of confirmed GMV, plus % of full-year GMV reached
--       as of each date — useful for a pacing-vs-target chart.
WITH running AS (
    SELECT
        sales_date,
        gmv_confirmed_vnd,
        SUM(gmv_confirmed_vnd) OVER (ORDER BY sales_date ROWS UNBOUNDED PRECEDING) AS ytd_gmv_vnd
    FROM shopee_sales_daily
)
SELECT
    sales_date,
    gmv_confirmed_vnd,
    ytd_gmv_vnd,
    ROUND(100.0 * ytd_gmv_vnd / SUM(gmv_confirmed_vnd) OVER (), 2) AS pct_of_annual_gmv_reached
FROM running
ORDER BY sales_date;

-- 14.2  Month-end YTD snapshot — the running total as of the last day of
--       each month (a compact version of 14.1 for a monthly pacing table).
WITH running AS (
    SELECT
        sales_date, sales_month,
        SUM(gmv_confirmed_vnd) OVER (ORDER BY sales_date ROWS UNBOUNDED PRECEDING) AS ytd_gmv_vnd,
        ROW_NUMBER() OVER (PARTITION BY sales_month ORDER BY sales_date DESC) AS rn_from_month_end
    FROM v_sales_daily
)
SELECT sales_month, sales_date AS month_end_date, ytd_gmv_vnd
FROM running
WHERE rn_from_month_end = 1
ORDER BY sales_month;


/* =============================================================================
   SECTION 15 — REPORTING VIEWS (for BI tool / dashboard consumption)
   -----------------------------------------------------------------------------
   These views package the recurring aggregations above into stable objects
   that a BI tool (Power BI, Tableau, Metabase, the project's own web
   dashboard, etc.) can connect to directly instead of re-deriving them.
   ============================================================================= */

-- 15.1  Monthly summary view — one row per month, every KPI the dashboard's
--       "Executive Overview" page needs.
CREATE OR REPLACE VIEW v_monthly_summary AS
SELECT
    sales_month,
    SUM(visitors)                                              AS visitors,
    SUM(buyers_placed)                                         AS buyers_placed,
    SUM(buyers_confirmed)                                      AS buyers_confirmed,
    SUM(gmv_placed_vnd)                                        AS gmv_placed_vnd,
    SUM(gmv_confirmed_vnd)                                     AS gmv_confirmed_vnd,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(visitors), 2)    AS conversion_rate_pct,
    ROUND(100.0 * SUM(buyers_confirmed) / SUM(buyers_placed), 2) AS confirmation_rate_pct,
    ROUND(SUM(gmv_confirmed_vnd) / SUM(buyers_confirmed), 0)   AS aov_vnd,
    ROUND(
        100.0 * (SUM(gmv_confirmed_vnd) - LAG(SUM(gmv_confirmed_vnd)) OVER (ORDER BY sales_month))
        / NULLIF(LAG(SUM(gmv_confirmed_vnd)) OVER (ORDER BY sales_month), 0), 2
    )                                                            AS mom_growth_pct
FROM v_sales_daily
GROUP BY sales_month;

-- 15.2  Campaign performance view — feeds the "Campaign Performance" dashboard
--       page (per-campaign-day GMV, uplift, and rank).
CREATE OR REPLACE VIEW v_campaign_performance AS
WITH month_avg AS (
    SELECT sales_month, AVG(gmv_confirmed_vnd) AS avg_gmv_vnd, AVG(visitors) AS avg_visitors
    FROM v_sales_daily
    GROUP BY sales_month
)
SELECT
    v.sales_date,
    v.sales_month,
    v.visitors,
    v.buyers_confirmed,
    v.gmv_confirmed_vnd,
    ROUND(v.gmv_confirmed_vnd / m.avg_gmv_vnd, 2) AS gmv_uplift_x,
    ROUND(v.visitors / m.avg_visitors, 2)          AS visitor_uplift_x,
    RANK() OVER (ORDER BY v.gmv_confirmed_vnd / m.avg_gmv_vnd DESC) AS uplift_rank
FROM v_sales_daily v
JOIN month_avg m ON m.sales_month = v.sales_month
WHERE v.is_campaign_day = 1;

-- 15.3  Conversion funnel view — one row per day, ready for the "Conversion
--       Funnel" dashboard page under any Month/Quarter/Campaign-day filter
--       combination (the dashboard aggregates this view client-side).
CREATE OR REPLACE VIEW v_funnel_daily AS
SELECT
    sales_date, sales_month, sales_quarter, half_year, is_campaign_day,
    visitors,
    buyers_placed, buyers_confirmed,
    units_placed, units_confirmed,
    orders_placed, orders_confirmed,
    gmv_placed_vnd, gmv_confirmed_vnd
FROM v_sales_daily;

-- 15.4  Day-of-week behavior view — feeds the "Customer Behavior" dashboard
--       page; keeps campaign-day rows tagged so the client can offer the
--       "exclude campaign days" toggle described in Section 7.
CREATE OR REPLACE VIEW v_day_of_week_behavior AS
SELECT
    day_of_week_num,
    day_of_week_name,
    is_campaign_day,
    COUNT(*)                          AS n_days,
    ROUND(AVG(gmv_confirmed_vnd), 0)  AS avg_gmv_vnd,
    ROUND(AVG(visitors), 0)           AS avg_visitors,
    ROUND(AVG(100.0 * buyers_confirmed / visitors), 2) AS avg_conversion_rate_pct
FROM v_sales_daily
GROUP BY day_of_week_num, day_of_week_name, is_campaign_day;

/* =============================================================================
   END OF FILE
   ============================================================================= */
