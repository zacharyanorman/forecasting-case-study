# Import libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 1) Load the three Excel sheets
sales_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Sales Opportunities")
revenue = pd.read_excel("case_study_data.xlsx", sheet_name="Revenue")
proj_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Projected Sales Opportunities")

# 2) Left join revenue with opportunities as new df
sales_revenue = revenue.merge(sales_opps, how="left", on="Opportunity ID")

# Make sure the important date columns are recognized as datetime objects
sales_revenue["Revenue Date"] = pd.to_datetime(sales_revenue["Revenue Date"])
sales_revenue["Enter Phase 2 Date"] = pd.to_datetime(sales_revenue["Enter Phase 2 Date"])
sales_opps["Open Date"] = pd.to_datetime(sales_opps["Opportunity Open Date"])
proj_opps["Sales Opportunity Month"] = pd.to_datetime(proj_opps["Sales Opportunity Month"])

# 3) Assign Phase 1 vs Phase 2

phase_list = []
for i in range(len(sales_revenue)):
    ph2_date = sales_revenue.loc[i, "Enter Phase 2 Date"]
    rev_date = sales_revenue.loc[i, "Revenue Date"]

    if pd.notna(ph2_date) and ph2_date <= rev_date:
        phase_list.append("Phase 2")
    else:
        phase_list.append("Phase 1")

sales_revenue["Phase"] = phase_list

# 4) Baseline metrics

# Get all opportunity IDs
all_opps = set(sales_opps["Opportunity ID"])

# Get converted IDs (those that appear in the revenue table)
converted_ids = set(sales_revenue["Opportunity ID"])

# Count domestic and international opportunities
domestic_opps = sales_opps.loc[sales_opps["Domestic"] == 1, "Opportunity ID"].nunique()
intl_opps = sales_opps.loc[sales_opps["International"] == 1, "Opportunity ID"].nunique()

# Count converted domestic and international opportunities
domestic_converted = sales_opps.loc[
    (sales_opps["Domestic"] == 1) & (sales_opps["Opportunity ID"].isin(converted_ids)), 
    "Opportunity ID"].nunique()

intl_converted = sales_opps.loc[
    (sales_opps["International"] == 1) & (sales_opps["Opportunity ID"].isin(converted_ids)),
    "Opportunity ID"
].nunique()

# Conversion rates
domestic_conv_rate = domestic_converted / domestic_opps if domestic_opps > 0 else 0
intl_conv_rate = intl_converted / intl_opps if intl_opps > 0 else 0

print(f"Domestic conversion rate: {domestic_conv_rate*100:.2f}%")
print(f"International conversion rate: {intl_conv_rate*100:.2f}%")

# Average sales events per converted opportunity
sales_counts_by_opp = sales_revenue.groupby("Opportunity ID")["Revenue"].size()
avg_sales_per_conversion = sales_counts_by_opp.mean()
print(f"Average sales events per converted opportunity: {avg_sales_per_conversion:.2f}")
print("\n")

# Average revenue per sale by region
avg_rev_per_sale_dom = sales_revenue.loc[sales_revenue["Domestic"] == 1, "Revenue"].mean()
avg_rev_per_sale_int = sales_revenue.loc[sales_revenue["International"] == 1, "Revenue"].mean()

print("Average revenue per sale:")
print(f"Domestic: ${avg_rev_per_sale_dom:,.2f}")
print(f"International: ${avg_rev_per_sale_int:,.2f}")
print("\n")

# 5) Build the projection using Projected Opportunities

# Step 1: expected sales per opportunity
exp_sales_per_opp_dom = domestic_conv_rate * avg_sales_per_conversion
exp_sales_per_opp_int = intl_conv_rate * avg_sales_per_conversion

# Step 2: estimate sales counts by multiplying opportunities × expected sales/opp
proj_opps["Domestic Prod 1 Sales"] = proj_opps["Domestic Product 1"] * exp_sales_per_opp_dom
proj_opps["Domestic Prod 2 Sales"] = proj_opps["Domestic Product 2"] * exp_sales_per_opp_dom
proj_opps["International Prod 1 Sales"] = proj_opps["International Product 1"] * exp_sales_per_opp_int
proj_opps["International Prod 2 Sales"] = proj_opps["International Product 2"] * exp_sales_per_opp_int

# Step 3: translate sales into revenue using avg revenue per sale
proj_opps["Domestic Prod 1 Rev"] = proj_opps["Domestic Prod 1 Sales"] * avg_rev_per_sale_dom
proj_opps["Domestic Prod 2 Rev"] = proj_opps["Domestic Prod 2 Sales"] * avg_rev_per_sale_dom
proj_opps["International Prod 1 Rev"] = proj_opps["International Prod 1 Sales"] * avg_rev_per_sale_int
proj_opps["International Prod 2 Rev"] = proj_opps["International Prod 2 Sales"] * avg_rev_per_sale_int

# Step 4: total projected revenue
proj_opps["Total Revenue"] = (
    proj_opps["Domestic Prod 1 Rev"]
  + proj_opps["Domestic Prod 2 Rev"]
  + proj_opps["International Prod 1 Rev"]
  + proj_opps["International Prod 2 Rev"]
)

# Step 5: aggregate monthly projection
proj_rev_m = proj_opps.groupby(pd.Grouper(key="Sales Opportunity Month", freq="MS"))["Total Revenue"].sum()
proj_rev_m.name = "Projection"

# 6) Historical revenue series
hist_rev_m = sales_revenue.groupby(pd.Grouper(key="Revenue Date", freq="MS"))["Revenue"].sum()
hist_rev_m.name = "History"
last_hist_month = hist_rev_m.dropna().index.max()

# 7) Combine and plot: History vs Projection
rev_df = pd.concat([hist_rev_m, proj_rev_m], axis=1)

plt.figure(figsize=(10, 5))
rev_df["History"].plot(label="History")
rev_df["Projection"].plot(label="Projection")
last_hist_month = hist_rev_m.dropna().index.max()
if pd.notna(last_hist_month):
    plt.axvline(last_hist_month, linestyle="--", color="gray", label="History Cutoff")
plt.title("Revenue: History vs Projection")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.legend()
plt.tight_layout()
plt.show()

# 8) Simple Organic Model

# Calculate average monthly growth rate from history (last 2 years)
monthly_changes = hist_rev_m.tail(24).pct_change()
avg_growth_rate = monthly_changes.tail(24).mean()
print(f"Average monthly growth rate (organic): {avg_growth_rate:.2%}")

# Forecast next 36 months (3 years) starting after the last history point
future_months = pd.date_range(last_hist_month + pd.offsets.MonthBegin(), periods=36, freq="MS")

# Start from the last actual historical revenue
last_val = hist_rev_m.dropna().iloc[-1] if len(hist_rev_m) > 0 else 0.0

organic_values = []
val = last_val
for i in range(len(future_months)):
    val = val * (1 + avg_growth_rate)
    organic_values.append(val)

organic_rev_m = pd.Series(organic_values, index=future_months, name="Organic")

# Combine into one dataframe (outer join keeps all dates)
rev_df = pd.concat([hist_rev_m, proj_rev_m, organic_rev_m], axis=1, join="outer")

# Plot all three 
plt.figure(figsize=(10, 5))
rev_df["History"].plot(label="History")
rev_df["Projection"].plot(label="Projection")
rev_df["Organic"].plot(label="Organic (Straight-line)", linestyle="--")
plt.axvline(last_hist_month, linestyle="--", color="gray", label="History Cutoff")
plt.title("Revenue: History vs Projection vs Organic Growth")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.legend()
plt.tight_layout()
plt.show()