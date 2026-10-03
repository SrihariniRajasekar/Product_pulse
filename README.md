# ProductPulse - E-commerce SKU Data Quality & Anomaly Dashboard

A data-quality pipeline for e-commerce SKU data. It profiles and cleans the records
(missing values, duplicate SKUs, invalid prices), stores them in SQLite, flags
pricing and stock anomalies using **IQR** and **z-score** outlier detection, and
exposes the results through an **Excel summary report** and a **Streamlit dashboard**.

**Stack:** Python, Pandas, NumPy, SQLite, SQL, Openpyxl, Streamlit, Altair

## Dataset
[Olist Brazilian E-Commerce](https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce)
(Kaggle, free). Tables used: `products`, `order_items`, `product_category_name_translation`.

## Quick start
```bash
pip install -r requirements.txt
python download_data.py      # needs a Kaggle login/token; or place the CSVs in data/raw manually
python pipeline.py           # clean -> SQLite -> anomalies -> Excel report
streamlit run app.py         # dashboard
```

## Pipeline
1. **Profile** raw data: missing values, duplicate SKUs/order lines, price <= 0,
   zero weight/dimensions, orphan order lines, SKUs with no sales -> `dq_issues` table.
2. **Clean:** dedupe SKUs and order lines, drop invalid-price and orphan lines,
   translate categories PT -> EN (fallback `unknown`), flag (not impute) invalid dimensions.
3. **Build SKU table:** units sold, median price (`current_price`), min/max/std price,
   average freight, number of sellers.
4. **Anomaly detection** (per category, minimum 20 SKUs, otherwise global stats):
   - *Pricing:* IQR fences (k = 1.5) and |z| > 3 on `log(price)`
   - *Stock:* IQR and z-score on `log1p(stock cover)` (upper tail = overstock)
   - *Rules:* negative stock; zero stock on top-5% sellers (stock-out risk)
   - *Severity:* 2 = both methods agree (or a hard rule broke), 1 = one method
5. **Store** in SQLite (`data/productpulse.db`): `sku_metrics`, `dq_issues`,
   `detection_eval`, `run_log`, plus SQL views `v_category_summary`, `v_anomalies`.
6. **Report:** `reports/productpulse_summary.xlsx` (Summary, Data Quality,
   Category Summary + chart, Top Anomalies, Detection Eval).

## Important assumptions (read before presenting this)
- **Olist has no inventory data.** `stock_level` is *simulated* from units sold with a
  fixed seed (`SEED = 42`), so results are reproducible. Real stock data would plug in at
  `add_synthetic_stock_and_inject()`.
- ~1% of SKUs each get **injected** anomalies (price spike, price drop, negative stock,
  overstock). `injected_anomaly` is the ground truth used only for the *Detection Eval*
  table (recall); the detectors never see it.
- `current_price` is the median selling price across all orders of that SKU.
- Prices are analysed on a log scale because they are heavily right-skewed.

## Useful SQL
```sql
SELECT * FROM v_category_summary ORDER BY anomaly_rate_pct DESC LIMIT 10;
SELECT * FROM v_anomalies WHERE severity = 2 LIMIT 50;
```

## Project layout
```
pipeline.py        cleaning, detection, SQLite, Excel report
app.py             Streamlit dashboard
download_data.py   Kaggle download helper
data/raw/          Olist CSVs (not committed)
data/productpulse.db
reports/productpulse_summary.xlsx
```
