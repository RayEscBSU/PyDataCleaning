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


""""
*** Cleaning Data ***
replace empty values with 0
drop unused columns, rename columns
split dist_num into 2 columns, one column stores the assessor number,
    the other column stores the district name
"""
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
# Check if the first digit of 'dist_nam' column is 0, if so, skip the first digit
df['assr_num'] = df['assr_num'].apply(lambda x: x[1:] if x.startswith('0') else x)
#convert assessor number column to int
df['assr_num'] = df['assr_num'].astype(int)
#remove digits that are now stored in assr_num col
df['dist_nam'] = df['dist_nam'].str[4:]

""" 
replace symbols with an empty string from dist_num

Since this data will be used for creating reports I
will leave these symbols in for ease of reading
"""
# # replace # 
# df['dist_nam'] = df['dist_nam'].str.replace(r'#', '')
# # replace / 
# df['dist_nam'] = df['dist_nam'].str.replace(r'/', '')
# # replace . 
# df['dist_nam'] = df['dist_nam'].str.replace(r'.', '')

#print(df.head())
#create a csv with results from changes
df.to_excel(r"L:\RayE\2024_County_Reports\PM006-Final_TaxPolicy_Cleaned.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')