import pandas as pd
df = pd.read_excel(r"C:\Users\rescobedo\Documents\test\Ada County (003).xls",sheet_name='Sheet1')
df = df.copy() #copy values only

#drop top tow rows
df = df.iloc[10:] # delete top rows
# Delete the last row
df = df.iloc[:-1]

#Cleaning Data
#replace Nan / NaT withe empty string
df = df.fillna('')
# Trim leading and trailing whitespaces from all string values
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

#df.to_excel(r"C:\Users\rescobedo\Documents\test\Ada County (003)draft.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')

df.insert(4, 'Acarage', df['Unnamed: 4'].astype(str) + df['Unnamed: 5'].astype(str) + df['DISTRICT ABSTRACT\nADA COUNTY\n8/1/2023 11:40:22AM'].astype(str))
df.insert(5, 'Market Value', df['Unnamed: 7'].astype(str) + df['Unnamed: 8'].astype(str) + df['Unnamed: 9'].astype(str))
df.insert(6, 'Homeowners Exception', df['Unnamed: 10'].astype(str) + df['Unnamed: 11'].astype(str))
df.insert(18, 'Ex', df['Unnamed: 16'].astype(str) + df['Unnamed: 18'].astype(str) + ' '+ df['Unnamed: 20'].astype(str))
df.insert(24, 'Net Value', df['Unnamed: 22'].astype(str) + df['Unnamed: 23'].astype(str) + df['Unnamed: 26'].astype(str))

#drop unwanted columns
df = df.drop(columns=['Unnamed: 4', 'Unnamed: 5', 'DISTRICT ABSTRACT\nADA COUNTY\n8/1/2023 11:40:22AM', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10',
                      'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16', 'Unnamed: 17', 'Unnamed: 18', 'Unnamed: 19', 'Unnamed: 20',
                      'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23','Unnamed: 24','Unnamed: 25', 'Unnamed: 26' ])
#Setting new column names
newColumnNames = ['Category Number', 'Unnamed: 1', ' Land Type 1', 'Land Type 2', 'Acarage','Market Value' , 'Homeowners Exception', 'Increment', 'Ex',
                  'Net Value']
df.columns = newColumnNames #assign new names to data frame

#need to create additional columns for all exceptions that we track and split the Ex into appropreate column 



#print
print(df.iloc[:10, :10])

#create a csv with results from changes
#df.to_excel(r"C:\Users\rescobedo\Documents\test\Ada County (003)TestOutput.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')