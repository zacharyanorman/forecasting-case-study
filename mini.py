# Import libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 1) Load the three Excel sheets
sales_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Sales Opportunities")
revenue = pd.read_excel("case_study_data.xlsx", sheet_name="Revenue")
proj_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Projected Sales Opportunities")

# 2) Merge revenue with opportunities and prepare date fields
sales_revenue = revenue.merge(sales_opps, how="left", on="Opportunity ID")

# Make sure the important date columns are recognized as datetime objects
sales_revenue["Revenue Date"] = pd.to_datetime(sales_revenue["Revenue Date"])
sales_revenue["Enter Phase 2 Date"] = pd.to_datetime(sales_revenue["Enter Phase 2 Date"])
sales_opps["Open Date"] = pd.to_datetime(sales_opps["Opportunity Open Date"])
proj_opps["Sales Opportunity Month"] = pd.to_datetime(proj_opps["Sales Opportunity Month"])

# 3) Assign Phase 1 vs Phase 2
# If Phase 2 date exists and is on/before revenue date → Phase 2
# Otherwise → Phase 1

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

# Average revenue per sale by region
avg_rev_per_sale_dom = sales_revenue.loc[sales_revenue["Domestic"] == 1, "Revenue"].mean()
avg_rev_per_sale_int = sales_revenue.loc[sales_revenue["International"] == 1, "Revenue"].mean()

print("Average revenue per sale:")
print(f"Domestic: ${avg_rev_per_sale_dom:,.2f}")
print(f"International: ${avg_rev_per_sale_int:,.2f}")

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

# Aggregate monthly projection
proj_rev_m = proj_opps.groupby(pd.Grouper(key="Sales Opportunity Month", freq="MS"))["Total Revenue"].sum()
proj_rev_m.name = "Projection"

# 6) Historical revenue series
hist_rev_m = sales_revenue.groupby(pd.Grouper(key="Revenue Date", freq="MS"))["Revenue"].sum()
hist_rev_m.name = "History"
last_hist_month = hist_rev_m.dropna().index.max()

# 7) Organic growth MVP

# Count historical opportunities by month
opps_hist = sales_opps.groupby(pd.Grouper(key="Open Date", freq="MS"))["Opportunity ID"].count()
opps_hist = opps_hist[opps_hist > 0]

# Calculate simple average monthly growth rate
monthly_changes = opps_hist.pct_change().dropna()
if len(monthly_changes) > 0:
    avg_growth_rate = monthly_changes.mean()
else:
    avg_growth_rate = 0.0

print(f"Organic avg monthly growth rate: {avg_growth_rate:.3%}")

# Forecast next 24 months of opportunities
future_months = pd.date_range(last_hist_month + pd.offsets.MonthBegin(), periods=24, freq="MS")
last_val = opps_hist.tail(12).mean() if len(opps_hist) > 0 else 0.0
if pd.isna(last_val):
    last_val = 0.0

organic_opps_list = []
val = last_val
for i in range(len(future_months)):
    val = val * (1 + avg_growth_rate)
    organic_opps_list.append(val)

organic_opps = pd.Series(organic_opps_list, index=future_months)

# Split organic opportunities into domestic vs international
domestic_share = sales_opps["Domestic"].mean()
intl_share = 1 - domestic_share

organic_dom_sales = organic_opps * domestic_share * exp_sales_per_opp_dom
organic_int_sales = organic_opps * intl_share * exp_sales_per_opp_int

organic_dom_rev = organic_dom_sales * avg_rev_per_sale_dom
organic_int_rev = organic_int_sales * avg_rev_per_sale_int

organic_rev = (organic_dom_rev + organic_int_rev).rename("Organic")

# 8) Combine and plot
rev_df = pd.concat([hist_rev_m, proj_rev_m, organic_rev], axis=1)

plt.figure(figsize=(10, 5))
rev_df["History"].plot(label="History")
rev_df["Projection"].plot(label="Projection")
rev_df["Organic"].plot(label="Organic (MVP)", linestyle="--")
if pd.notna(last_hist_month):
    plt.axvline(last_hist_month, linestyle="--", color="gray", label="History Cutoff")
plt.title("Revenue: History vs Projection vs Organic")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.legend()
plt.tight_layout()
plt.show()

# 9) Print summary stats for narration
print("\nSummary")
print(f"Domestic conversion rate: {domestic_conv_rate:.3f}")
print(f"International conversion rate: {intl_conv_rate:.3f}")
print(f"Avg sales per converted opp: {avg_sales_per_conversion:.2f}")
print(f"Avg revenue per sale Domestic: {avg_rev_per_sale_dom:,.2f}")
print(f"Avg revenue per sale International: {avg_rev_per_sale_int:,.2f}")
print(f"Organic MVP avg monthly growth rate: {avg_growth_rate:.3%}")