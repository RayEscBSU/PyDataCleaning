import pandas as pd
df = pd.read_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2023 Example Reports\GenTaxYR045.xls")

df = df.copy() #copy values only

#print(df.columns)
#Setting new column names
newColumnNames = ['County Number', 'Unnamed: 1', 'Unnamed: 2', 'County', 'Unnamed: 4',
       'Unnamed: 5', 'Unnamed: 6', 'DistrictType', 'DistrictNumber', 'DistrictName',
       'RAA', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13',
       'Unnamed: 14', 'Unnamed: 15', 'TaxableValue', 'Unnamed: 17',
       'BaseValue', 'Unnamed: 19', 'Unnamed: 20', 'IncrementValue',
       'Unnamed: 22', 'Unnamed: 23', 'AdjustedTaxableValue', 'Unnamed: 25']

df.columns = newColumnNames #assign new names to data frame

#drop unwanted columns
df = df.drop(columns=[ 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 4',
       'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13',
       'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 17', 'Unnamed: 19', 'Unnamed: 20',
       'Unnamed: 22', 'Unnamed: 23', 'Unnamed: 25'])
#df.columns = newColumnNames #assign new names to data frame

#drop top tow rows
df = df.iloc[16:] # delete top rows
# Delete the last row
df = df.iloc[:-7]

"Align last for column values to match district name"
df.reset_index(drop=True, inplace=True)
# Select the last four columns
last_four_columns = df.iloc[:, -4:]
# Remove the specified row (index 1) and shift data up
last_four_columns = last_four_columns.drop(index=1).reset_index(drop=True)
# Keep the first part of the DataFrame (excluding the last four columns)
first_columns = df.iloc[:, :-4]
# Adjust the length of the first part to match the new length of the last part
first_columns = first_columns.iloc[:len(last_four_columns), :]
# Combine the first columns with the modified last four columns
df = pd.concat([first_columns, last_four_columns], axis=1)

# Add 2 rows to the first column
df['County'] = pd.concat([pd.Series([None, None]), df['County'].reset_index(drop=True)], ignore_index=True)

# Add 1 row to the second column
df['DistrictType'] = pd.concat([pd.Series([None]), df['DistrictType'].reset_index(drop=True)], ignore_index=True)
df = df.iloc[2:] # delete top rows

# Delete any rows that contain only 0 or empty string
df = df.dropna(how='all')

# Extract the digits and county names
df['County Number'] = df['County'].str.extract('(\d+)').astype(float).astype('Int64')
df['County'] = df['County'].str.extract('- (.+)')[0]

# Remove leading zero from County Number
df['County Number'] = df['County Number'].apply(lambda x: str(x).lstrip('0') if pd.notnull(x) else x)


#fill empty cells in column with last value seen in column
df[['County Number', 'County', 'DistrictType', 'DistrictNumber']] = df[['County Number', 'County', 'DistrictType', 'DistrictNumber']].fillna(method='ffill')
#replace Nan / NaT withe empty string
df.iloc[:, -4:] = df.iloc[:, -4:].fillna(0)


print(df.iloc[:10, :10])
#create a csv with results from changes
#df.to_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2023 Example Reports\GenTaxYR045Modified.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')
df.to_csv(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2023 Example Reports\GenTaxYR045Modified.csv", index=False, header=True)
