"""
Aerospace & Industrial Automation Market Research Analysis
Temple Allen Industries – Business Development Portfolio
Author: Yadnesh Deshpande
"""

import pandas as pd
import numpy as np
import json

# ── Market Data ──────────────────────────────────────────────────────────────

market_data = {
    "Segment": ["Aerospace MRO", "Naval/Marine", "Wind Energy", "Transportation", "Defense"],
    "Market_Size_2023_USD_B": [81.9, 34.2, 22.7, 18.4, 15.6],
    "CAGR_Percent": [5.8, 4.2, 9.1, 6.3, 3.9],
    "Automation_Adoption_Percent": [38, 22, 45, 31, 55],
    "Key_Pain_Point": [
        "Manual sanding/surface prep is slow & hazardous",
        "Corrosion treatment labor-intensive",
        "Blade surface finishing inconsistent",
        "Paint prep compliance issues",
        "Strict surface quality requirements"
    ],
    "EMMA_Fit_Score": [10, 8, 7, 6, 9]  # Out of 10
}

df = pd.DataFrame(market_data)

# ── Projections ───────────────────────────────────────────────────────────────

years = [2024, 2025, 2026, 2027, 2028]
projections = []

for _, row in df.iterrows():
    base = row["Market_Size_2023_USD_B"]
    cagr = row["CAGR_Percent"] / 100
    for yr in years:
        n = yr - 2023
        projected = round(base * ((1 + cagr) ** n), 2)
        projections.append({
            "Segment": row["Segment"],
            "Year": yr,
            "Projected_Market_Size_USD_B": projected
        })

proj_df = pd.DataFrame(projections)

# ── Opportunity Sizing ────────────────────────────────────────────────────────

df["Addressable_Market_2028_USD_B"] = df.apply(
    lambda r: round(
        r["Market_Size_2023_USD_B"]
        * ((1 + r["CAGR_Percent"] / 100) ** 5)
        * (r["Automation_Adoption_Percent"] / 100),
        2
    ),
    axis=1
)

df["TAI_Opportunity_Score"] = (
    df["EMMA_Fit_Score"] * 0.5 +
    df["Automation_Adoption_Percent"] / 10 * 0.3 +
    df["CAGR_Percent"] / 10 * 0.2
).round(2)

# ── Competitive Landscape ─────────────────────────────────────────────────────

competitors = {
    "Company": ["Electroimpact", "Kuka Aerospace", "Dürr AG", "Temple Allen (EMMA)", "Genesis Systems"],
    "Automation_Type": ["Drilling/Fastening", "Assembly", "Surface Finishing", "Sanding/Prep", "Welding"],
    "Target_Market": ["Aerospace", "Aerospace", "Auto/Aero", "Aero/Marine/Wind", "General Mfg"],
    "Avg_System_Price_USD_K": [850, 1200, 600, 350, 280],
    "AI_ML_Integrated": [False, False, True, True, False],
    "Operates_Alongside_Humans": [False, False, False, True, False]
}

comp_df = pd.DataFrame(competitors)

# ── Summary Output ────────────────────────────────────────────────────────────

print("=" * 60)
print("MARKET OPPORTUNITY SUMMARY – Temple Allen Industries")
print("=" * 60)

print("\n[1] Market Segments by EMMA Fit Score:")
print(df[["Segment", "Market_Size_2023_USD_B", "CAGR_Percent",
          "Addressable_Market_2028_USD_B", "TAI_Opportunity_Score"]]
      .sort_values("TAI_Opportunity_Score", ascending=False)
      .to_string(index=False))

print("\n[2] Top Opportunity: Aerospace MRO")
top = df[df["Segment"] == "Aerospace MRO"].iloc[0]
print(f"   Market Size 2023:       ${top['Market_Size_2023_USD_B']}B")
print(f"   CAGR:                   {top['CAGR_Percent']}%")
print(f"   Automation Adoption:    {top['Automation_Adoption_Percent']}%")
print(f"   Addressable Mkt 2028:   ${top['Addressable_Market_2028_USD_B']}B")

print("\n[3] Competitive Positioning – EMMA vs Competitors:")
print(comp_df.to_string(index=False))

print("\n[4] EMMA Differentiators:")
print("   ✓ Only system that operates alongside humans (no cage/exclusion zone)")
print("   ✓ Real-time ML & computer vision for adaptive sanding")
print("   ✓ Lowest price point vs comparable automation ($350K vs $600K–$1.2M)")
print("   ✓ No infrastructure changes required for deployment")

# ── Save outputs ──────────────────────────────────────────────────────────────

df.to_csv("market_segments.csv", index=False)
proj_df.to_csv("market_projections.csv", index=False)
comp_df.to_csv("competitive_landscape.csv", index=False)

print("\n✓ CSVs saved: market_segments.csv, market_projections.csv, competitive_landscape.csv")
