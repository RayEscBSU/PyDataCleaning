import pandas as pd

# Set option to opt-in to future behavior
pd.set_option('future.no_silent_downcasting', True)

#reac excel file
df = pd.read_excel(r"L:\RayE\2024_County_Reports\PM005-Final_Tax Policy.xls")

print(df.head())
#create a csv with results from changes
#df.to_excel(r"L:\RayE\2024_County_Reports\PM005-Final_TaxPolicy_Cleaned.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')