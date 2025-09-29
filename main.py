import pandas as pd

# Import sales and revenue sheets
sales_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Sales Opportunities")
revenue = pd.read_excel("case_study_data.xlsx", sheet_name="Revenue")

# Merge sales and revenue
sales_revenue = revenue.merge(sales_opps, how='left',on='Opportunity ID')

# Create new column "Phase"
sales_revenue["Phase"] = ""

# Add either "Phase 1" or "Phase 2" to "Phase" column, 
# depending on when sale occured 

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

# Domestic Phase 1
domestic_ph1_sales = sales_revenue[(sales_revenue["Domestic"] == 1) & (sales_revenue["Phase"] == "Phase 1")]
avg_domestic_ph1_rev = domestic_ph1_sales["Revenue"].mean()

# Domestic Phase 2
domestic_ph2_sales = sales_revenue[(sales_revenue["Domestic"] == 1) & (sales_revenue["Phase"] == "Phase 2")]
avg_domestic_ph2_rev = domestic_ph2_sales["Revenue"].mean()

# International Phase 1
intl_ph1_sales = sales_revenue[(sales_revenue["International"] == 1) & (sales_revenue["Phase"] == "Phase 1")]
avg_intl_ph1_rev = intl_ph1_sales["Revenue"].mean()

# International Phase 2
intl_ph2_sales = sales_revenue[(sales_revenue["International"] == 1) & (sales_revenue["Phase"] == "Phase 2")]
avg_intl_ph2_rev = intl_ph2_sales["Revenue"].mean()

print("Average revenue per sale:")
print(f"Domestic Phase 1: ${avg_domestic_ph1_rev:.2f}")
print(f"Domestic Phase 2: ${avg_domestic_ph2_rev:.2f}")
print(f"International Phase 1: ${avg_intl_ph1_rev:.2f}")
print(f"International Phase 2: ${avg_intl_ph2_rev:.2f}")