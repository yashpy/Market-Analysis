"""
Automated ETL Pipeline: Excel/CSV → SQL Database
Temple Allen Industries – Business Development Portfolio
Author: Yadnesh Deshpande

Workflow:
  1. Read equipment & orders data from Excel/CSV
  2. Clean & validate data
  3. Load into SQLite (mirrors PostgreSQL logic)
  4. Generate summary report
"""

import pandas as pd
import sqlite3
import os
from datetime import datetime

DB_PATH = "tai_equipment.db"

# ── Step 1: Generate Sample Input Data ───────────────────────────────────────

units_data = {
    "serial_number":   ["EMMA-2024-001","EMMA-2024-002","EMMA-2024-003","EMMA-2024-004","EMMA-2025-001"],
    "model_version":   ["v2.1","v2.1","v2.2","v2.2","v2.3"],
    "customer":        ["Boeing MRO","Lockheed Martin","Vestas Wind","Huntington Ingalls",""],
    "industry":        ["Aerospace","Defense","Windpower","Marine",""],
    "status":          ["Deployed","Deployed","Demo","Deployed","In Stock"],
    "deployed_date":   ["2024-03-01","2024-04-15","2024-06-01","2024-08-20",""],
    "unit_price_usd":  [350000, 350000, 0, 350000, 350000],
}

orders_data = {
    "serial_number":   ["EMMA-2024-001","EMMA-2024-001","EMMA-2024-002","EMMA-2024-003","EMMA-2024-004"],
    "item_name":       ["Sanding Disc 80-grit","Backing Pad","Sanding Disc 120-grit","Dust Filter","Sanding Disc 220-grit"],
    "quantity":        [200, 5, 150, 3, 100],
    "order_date":      ["2024-06-01","2024-07-15","2024-07-01","2024-08-01","2024-09-10"],
    "unit_price":      [2.50, 18.00, 2.75, 45.00, 3.00],
}

units_df  = pd.DataFrame(units_data)
orders_df = pd.DataFrame(orders_data)

# Save as Excel (simulates real input source)
with pd.ExcelWriter("tai_field_data.xlsx", engine="openpyxl") as writer:
    units_df.to_excel(writer,  sheet_name="Units",  index=False)
    orders_df.to_excel(writer, sheet_name="Orders", index=False)

print("✓ Input Excel created: tai_field_data.xlsx")

# ── Step 2: Read & Clean ──────────────────────────────────────────────────────

xl = pd.read_excel("tai_field_data.xlsx", sheet_name=None)
units  = xl["Units"].copy()
orders = xl["Orders"].copy()

# Clean units
units["serial_number"] = units["serial_number"].str.strip().str.upper()
units["customer"]      = units["customer"].fillna("").str.strip()
units["industry"]      = units["industry"].fillna("Unknown").str.strip()
units["status"]        = units["status"].str.strip()
units["deployed_date"] = pd.to_datetime(units["deployed_date"], errors="coerce")
units["is_deployed"]   = units["status"] == "Deployed"
units["loaded_at"]     = datetime.now()

# Clean orders
orders["serial_number"] = orders["serial_number"].str.strip().str.upper()
orders["order_date"]    = pd.to_datetime(orders["order_date"])
orders["total_cost"]    = orders["quantity"] * orders["unit_price"]
orders["loaded_at"]     = datetime.now()

# Validate: flag orphan orders
valid_serials = set(units["serial_number"])
orphans = orders[~orders["serial_number"].isin(valid_serials)]
if not orphans.empty:
    print(f"⚠ Warning: {len(orphans)} orders reference unknown unit serial(s): {orphans['serial_number'].tolist()}")

print(f"✓ Cleaned: {len(units)} units, {len(orders)} orders")

# ── Step 3: Load to Database ──────────────────────────────────────────────────

conn = sqlite3.connect(DB_PATH)

units.to_sql("emma_units",       conn, if_exists="replace", index=False)
orders.to_sql("consumable_orders", conn, if_exists="replace", index=False)

print(f"✓ Loaded to database: {DB_PATH}")

# ── Step 4: Summary Queries ───────────────────────────────────────────────────

print("\n" + "="*55)
print("TAI FIELD EQUIPMENT REPORT")
print("="*55)

# Units by status
status_q = pd.read_sql("""
    SELECT status, COUNT(*) as count
    FROM emma_units
    GROUP BY status ORDER BY count DESC
""", conn)
print("\n[1] Units by Status:")
print(status_q.to_string(index=False))

# Units by industry
industry_q = pd.read_sql("""
    SELECT industry,
           COUNT(*) as units,
           SUM(unit_price_usd) as total_revenue
    FROM emma_units
    WHERE status = 'Deployed'
    GROUP BY industry ORDER BY total_revenue DESC
""", conn)
print("\n[2] Deployed Units by Industry:")
print(industry_q.to_string(index=False))

# Top consumable orders
orders_q = pd.read_sql("""
    SELECT item_name,
           SUM(quantity) as total_qty,
           ROUND(SUM(total_cost), 2) as total_spend
    FROM consumable_orders
    GROUP BY item_name ORDER BY total_spend DESC
""", conn)
print("\n[3] Consumables Spend:")
print(orders_q.to_string(index=False))

# Total revenue
rev_q = pd.read_sql("""
    SELECT
        COUNT(*) as total_units,
        SUM(CASE WHEN status='Deployed' THEN unit_price_usd ELSE 0 END) as deployed_revenue,
        SUM(unit_price_usd) as total_pipeline
    FROM emma_units
""", conn)
print("\n[4] Revenue Summary:")
print(rev_q.to_string(index=False))

conn.close()

# Save summary to CSV
orders_q.to_csv("consumables_summary.csv", index=False)
print("\n✓ Summary saved: consumables_summary.csv")
print("✓ Pipeline complete.")
