import pandas as pd

# Set option to opt-in to future behavior
pd.set_option('future.no_silent_downcasting', True)

#reac excel file
df = pd.read_excel(r"L:\RayE\2024_County_Reports\PM006-Final_TaxPolicy.xls")

#save new header 
new_header = df.iloc[1]
#drop top tow rows
df = df.iloc[3:] # delete top rows
# Delete the last row
df = df.iloc[:-2]

#set new header 
df.columns = new_header


#Cleaning Data
#replace Nan / NaT withe empty string
df.fillna(0, inplace =True)
# Trim leading and trailing whitespaces from all string values
#df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

#print col names
#print(df.columns)

#drop unused colums
df=df.drop(columns=['PP EXEMPTION  (5)'])
#rename columns
newColumnNames = ['dist_nam', 'real_p_v', 'est_sub', 'main_roll_h_e','annex' , 'incr_val']
df.columns = newColumnNames #assign new names to data frame

#Remove rows where 'Taxing District (1)' does not contain a number
df = df[df['dist_nam'].str.contains('\d', na=False)]

# Insert new column at position 0 containing the substring of the first three characters from the 'dist_nam' column
df.insert(0, 'assr_num', df['dist_nam'].str[:3])

# Remove first 4 characters from 'dist_nam' column
df['dist_nam'] = df['dist_nam'].str[4:]

print(df.head())
#create a csv with results from changes
df.to_excel(r"L:\RayE\2024_County_Reports\PM006-Final_TaxPolicy_Cleaned.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')