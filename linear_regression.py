import pandas as pd
import statsmodels.api as sm

sales_opps = pd.read_excel("case_study_data.xlsx", sheet_name="Sales Opportunities")
revenue = pd.read_excel("case_study_data.xlsx", sheet_name="Revenue")

sales_revenue = revenue.merge(sales_opps, how='left', on='Opportunity ID')

X = sales_revenue[["Product 1", "International", "High Potential", "PH2", "Active"]]
X = sm.add_constant(X)

model = sm.OLS(sales_revenue["Revenue"], X).fit()
print(model.summary())
