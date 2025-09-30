# Import libraries
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Load sales and revenue sheets as variables
sales_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Sales Opportunities")
revenue = pd.read_excel("case_study_data.xlsx", sheet_name="Revenue")

# Merge/join sales and revenue
sales_revenue = revenue.merge(sales_opps, how='left', on='Opportunity ID')

# Create new column "Phase"
sales_revenue["Phase"] = ""

# Add either "Phase 1" or "Phase 2" to "Phase" column, depending on when sale occured 
for i in range(len(sales_revenue)):
    rev_date = sales_revenue.loc[i, "Revenue Date"]
    ph2_date = sales_revenue.loc[i, "Enter Phase 2 Date"]

    if pd.isna(ph2_date):
        sales_revenue.loc[i, "Phase"] = "Phase 1"
    elif rev_date < ph2_date:
        sales_revenue.loc[i, "Phase"] = "Phase 1"
    elif ph2_date <= rev_date:
        sales_revenue.loc[i, "Phase"] = "Phase 2"

# Print preview to confirm all columns are present and contain correct data
print(sales_revenue.head())
print(sales_revenue.columns)

# Calculate baseline metrics:

# Overall conversion rate
total_opps = len(sales_opps)
converted_opps = sales_revenue["Opportunity ID"].nunique()

conversion_rate = converted_opps/total_opps
print(f"The historic sale conversion rate is: {(conversion_rate*100):.2f}%")

# PH1 and PH2 conversion rates
ph1_sales_uniq_opp_id = sales_revenue.loc[sales_revenue["Phase"] == "Phase 1", "Opportunity ID"].nunique()
ph1_conversions = ph1_sales_uniq_opp_id / total_opps

ph2_sales_uniq_opp_id = sales_revenue.loc[sales_revenue["Phase"] == "Phase 2", "Opportunity ID"].nunique()
ph2_conversions = ph2_sales_uniq_opp_id / total_opps

print(f"The historic PH1 sale conversion rate is: {(ph1_conversions*100):.2f}%")
print(f"The historic PH2 sale conversion rate is: {(ph2_conversions*100):.2f}%")

# Average PH1 and PH2 revenue
ph1_total_rev = 0
ph2_total_rev = 0

ph1_total_sales_events = 0
ph2_total_sales_events = 0

for i in range(len(sales_revenue)):
    if sales_revenue.loc[i, "Phase"] == "Phase 1":
        ph1_total_rev += sales_revenue.loc[i, "Revenue"]
        ph1_total_sales_events += 1
    if sales_revenue.loc[i, "Phase"] == "Phase 2":
        ph2_total_rev += sales_revenue.loc[i, "Revenue"]
        ph2_total_sales_events +=1

avg_ph1_rev_per_sale = ph1_total_rev/ph1_total_sales_events
avg_ph2_rev_per_sale = ph2_total_rev/ph2_total_sales_events

print(f"The average revenue for a PH1 sale is: ${avg_ph1_rev_per_sale:.2f}")
print(f"The average revenue for a PH2 sale is: ${avg_ph2_rev_per_sale:.2f}")

# Average number of sales per conversion
avg_ph1_sales_per_conversion = ph1_total_sales_events/ph1_sales_uniq_opp_id
avg_ph2_sales_per_conversion = ph2_total_sales_events/ph2_sales_uniq_opp_id

print(f"The avg. number of sales for a converted PH1 opportunity is: {avg_ph1_sales_per_conversion:.2f}")
print(f"The avg. number of sales for a converted PH2 opportunity is: {avg_ph2_sales_per_conversion:.2f}")

# Domestic vs international conversion rates
domestic_opps = sales_opps[sales_opps["Domestic"] == 1]["Opportunity ID"].nunique()
intl_opps = sales_opps[sales_opps["International"] == 1]["Opportunity ID"].nunique()

domestic_converted = sales_revenue[sales_revenue["Domestic"] == 1]["Opportunity ID"].nunique()
intl_converted = sales_revenue[sales_revenue["International"] == 1]["Opportunity ID"].nunique()

domestic_conv_rate = domestic_converted / domestic_opps
intl_conv_rate = intl_converted / intl_opps

print(f"The historic domestic conversion rate is: {(domestic_conv_rate*100):.2f}%")
print(f"The historic international conversion rate is: {(intl_conv_rate*100):.2f}%")

# Domestic vs international revenues
domestic_ph1_sales = sales_revenue[(sales_revenue["Domestic"] == 1) & (sales_revenue["Phase"] == "Phase 1")]
avg_domestic_ph1_rev = domestic_ph1_sales["Revenue"].mean()

domestic_ph2_sales = sales_revenue[(sales_revenue["Domestic"] == 1) & (sales_revenue["Phase"] == "Phase 2")]
avg_domestic_ph2_rev = domestic_ph2_sales["Revenue"].mean()

intl_ph1_sales = sales_revenue[(sales_revenue["International"] == 1) & (sales_revenue["Phase"] == "Phase 1")]
avg_intl_ph1_rev = intl_ph1_sales["Revenue"].mean()

intl_ph2_sales = sales_revenue[(sales_revenue["International"] == 1) & (sales_revenue["Phase"] == "Phase 2")]
avg_intl_ph2_rev = intl_ph2_sales["Revenue"].mean()

print("Average revenue per sale:")
print(f"Domestic Phase 1: ${avg_domestic_ph1_rev:.2f}")
print(f"Domestic Phase 2: ${avg_domestic_ph2_rev:.2f}")
print(f"International Phase 1: ${avg_intl_ph1_rev:.2f}")
print(f"International Phase 2: ${avg_intl_ph2_rev:.2f}")

# Create the model for projected growth:

# Access the Project Opportunities sheet
proj_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Projected Sales Opportunities")

# Create sales columns
exp_sales_per_opp_dom = (domestic_conv_rate * avg_ph1_sales_per_conversion) + (domestic_conv_rate * avg_ph2_sales_per_conversion)
exp_sales_per_opp_intl = (intl_conv_rate * avg_ph1_sales_per_conversion) + (intl_conv_rate * avg_ph2_sales_per_conversion)

proj_opps["Domestic Prod 1 Sales"] = exp_sales_per_opp_dom * proj_opps["Domestic Product 1"]
proj_opps["Domestic Prod 2 Sales"] = exp_sales_per_opp_dom * proj_opps["Domestic Product 2"]
proj_opps["International Prod 1 Sales"] = exp_sales_per_opp_intl * proj_opps["International Product 1"]
proj_opps["International Prod 2 Sales"] = exp_sales_per_opp_intl * proj_opps["International Product 2"]

# Create revenue columns
exp_rev_per_sale_dom = ((avg_domestic_ph1_rev * avg_ph1_sales_per_conversion)
                        + (avg_domestic_ph2_rev * avg_ph2_sales_per_conversion)
                        ) / (avg_ph1_sales_per_conversion + avg_ph2_sales_per_conversion)

exp_rev_per_sale_int = ((avg_intl_ph1_rev * avg_ph1_sales_per_conversion)
                        + (avg_intl_ph2_rev * avg_ph2_sales_per_conversion)
                        ) / (avg_ph1_sales_per_conversion + avg_ph2_sales_per_conversion)

proj_opps["Domestic Prod 1 Rev"] = proj_opps["Domestic Prod 1 Sales"] * exp_rev_per_sale_dom
proj_opps["Domestic Prod 2 Rev"] = proj_opps["Domestic Prod 2 Sales"] * exp_rev_per_sale_dom
proj_opps["International Prod 1 Rev"] = proj_opps["International Prod 1 Sales"] * exp_rev_per_sale_int
proj_opps["International Prod 2 Rev"] = proj_opps["International Prod 2 Sales"] * exp_rev_per_sale_int

# Create Total Sales and Total Revenue
proj_opps["Total Sales"] = (proj_opps["Domestic Prod 1 Sales"] + proj_opps["Domestic Prod 2 Sales"] + 
                            proj_opps["International Prod 1 Sales"] + proj_opps["International Prod 2 Sales"])

proj_opps["Total Revenue"] = (proj_opps["Domestic Prod 1 Rev"] + proj_opps["Domestic Prod 2 Rev"]+ 
                              proj_opps["International Prod 1 Rev"] + proj_opps["International Prod 2 Rev"])

# Print preview to confirm all columns are present and contain correct data
print(proj_opps.head())
print(proj_opps.columns)

# Forecast Revenue based on 'Projected Sales Opportunities'

# Convert sales and opps to datetime
sales_revenue["Revenue Date"] = pd.to_datetime(sales_revenue["Revenue Date"])
proj_opps["Sales Opportunity Month"] = pd.to_datetime(proj_opps["Sales Opportunity Month"])

# Ensure both 'History' and 'Projection' are in month format
hist_rev_m = (sales_revenue.groupby(pd.Grouper(key="Revenue Date", freq="MS"))["Revenue"].sum().rename("History"))
proj_rev_m = (proj_opps.groupby(pd.Grouper(key="Sales Opportunity Month", freq="MS"))["Total Revenue"].sum().rename("Projection"))

# Combine history and projection for plotting
rev_df = pd.concat([hist_rev_m, proj_rev_m], axis=1)

# Plot History and projection
plt.figure(figsize=(10,5))
rev_df["History"].plot()
rev_df["Projection"].plot()

# Create vertical line at the last actual month
last_hist = hist_rev_m.dropna().index.max()
if pd.notna(last_hist):
    plt.axvline(last_hist, linestyle="--")

# Label and show history and projection
plt.title("Revenue: History vs Projection (Monthly)")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.legend()
plt.tight_layout()
plt.show()

# Plot projected opportunies as a check for dip

# Convert to Datetime
proj_opps["Sales Opportunity Month"] = pd.to_datetime(proj_opps["Sales Opportunity Month"])

# Total opportunities per month (sum across all 4 products)
proj_opps["Total Opps"] = (proj_opps["Domestic Product 1"] + proj_opps["Domestic Product 2"]
                           + proj_opps["International Product 1"]+ proj_opps["International Product 2"])

# Plot opportunities and show plot
plt.figure(figsize=(10,5))
plt.plot(proj_opps["Sales Opportunity Month"], proj_opps["Total Opps"], label="Projected Opportunities")

plt.title("Projected Opportunities Over Time")
plt.xlabel("Month")
plt.ylabel("Total Opportunities")
plt.legend()
plt.tight_layout()
plt.show()

# Forecast 5 years of Sales and Revenue, where sales opportunity volume grows organically

# Create 'Open Date' column filled with historical opportunities data, removing zeros
sales_opps["Open Date"] = pd.to_datetime(sales_opps["Opportunity Open Date"])
opps_hist = sales_opps.groupby(pd.Grouper(key="Open Date", freq="MS"))["Opportunity ID"].count()
opps_hist = opps_hist[opps_hist > 0]

# Calculate monthly % changes across full history, dropping missing values
r = opps_hist.pct_change().dropna()

# Calculate geometric mean growth rate (in lieu of arithmetic mean)
avg_growth_rate = np.exp(np.log1p(r).mean()) - 1
print(f"The average geometric growth rate is: {(avg_growth_rate * 100):.2f}%")

# Populate future months for next five years
future_months = pd.date_range("2016-08-01", periods=60, freq="MS")

# Create last/basis of growth factor based on last 12 months of growth
last_val = opps_hist.tail(12).mean()

# Create empty list of organic opportunities 
organic_opps = []

# Loop through each future month and apply growth
for i in future_months:
    last_val = last_val * (1 + avg_growth_rate) 
    organic_opps.append(last_val)                

# Turn the list into a series alongside future months
organic_opps = pd.Series(organic_opps, index=future_months)

# Calculate (share of) domestic and international opportunities
domestic_share = sales_opps["Domestic"].mean()
intl_share = 1 - domestic_share

organic_domestic_opps = organic_opps * domestic_share
organic_international_opps = organic_opps * intl_share

# Calculated domestic and international sales
organic_domestic_sales = organic_domestic_opps * exp_sales_per_opp_dom
organic_international_sales = organic_international_opps * exp_sales_per_opp_intl

# Calculate total organic sales and revenue
organic_domestic_revenue = organic_domestic_sales * exp_rev_per_sale_dom
organic_international_revenue = organic_international_sales * exp_rev_per_sale_int

organic_sales_total = organic_domestic_sales + organic_international_sales
organic_revenue_total = organic_domestic_revenue + organic_international_revenue

# Name column before adding to df
organic_rev_m = organic_revenue_total.rename("Organic")

# Create df to plot from
rev_compare = pd.DataFrame({"History":   hist_rev_m,
                            "Projected": proj_rev_m,
                            "Organic":   organic_rev_m})

# Plot the hisoric, projected, and organic revenue
plt.figure(figsize=(10,5))
rev_compare["History"].plot(label="History")
rev_compare["Projected"].plot(label="Projected")
rev_compare["Organic"].plot(label="Organic")
last_hist = hist_rev_m.dropna().index.max()

# Add line where historical data ends
plt.axvline(last_hist, linestyle="--", color="gray", label="History Cutoff")

# Label and show Revenue Forecast
plt.title("Revenue Forecast: History vs Projection vs Organic")
plt.xlabel("Month")
plt.ylabel("Revenue")
plt.legend()
plt.tight_layout()
plt.show()