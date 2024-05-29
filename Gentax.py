import pandas as pd
import sqlalchemy as db
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

# Load the Excel file
df = pd.read_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2023 Example Reports\GenTaxYR045.xls")

df = df.copy()  # Copy values only

# Setting new column names
newColumnNames = ['CountyNumber', 'Unnamed: 1', 'Unnamed: 2', 'County', 'Unnamed: 4',
                  'Unnamed: 5', 'Unnamed: 6', 'DistrictType', 'DistrictNumber', 'DistrictName',
                  'RAA', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13',
                  'Unnamed: 14', 'Unnamed: 15', 'TaxableValue', 'Unnamed: 17',
                  'BaseValue', 'Unnamed: 19', 'Unnamed: 20', 'IncrementValue',
                  'Unnamed: 22', 'Unnamed: 23', 'AdjustedTaxableValue', 'Unnamed: 25']

df.columns = newColumnNames  # Assign new names to data frame

# Drop unwanted columns
df = df.drop(columns=['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 4',
                      'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13',
                      'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 17', 'Unnamed: 19', 'Unnamed: 20',
                      'Unnamed: 22', 'Unnamed: 23', 'Unnamed: 25'])

# Drop top rows
df = df.iloc[16:]  # Delete top rows
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
df = df.iloc[2:]  # delete top rows

# Delete any rows that contain only NaN values
df = df.dropna(how='all')

# Extract the digits and county names
df['CountyNumber'] = df['County'].str.extract(r'(\d+)').astype(float).astype('Int64')
df['County'] = df['County'].str.extract(r'- (.+)')[0]

# Convert County Number to integers and remove leading zeros
df['CountyNumber'] = df['CountyNumber'].apply(lambda x: int(x) if pd.notnull(x) else x)

# Convert DistrictType and DistrictNumber to integers
df['DistrictType'] = df['DistrictType'].apply(lambda x: int(x) if pd.notnull(x) else x)
df['DistrictNumber'] = df['DistrictNumber'].apply(lambda x: int(x) if pd.notnull(x) else x)

# Forward fill empty cells in specified columns
df[['CountyNumber', 'County', 'DistrictType', 'DistrictNumber', 'DistrictName']] = df[['CountyNumber', 'County', 'DistrictType', 'DistrictNumber', 'DistrictName']].ffill()

# Replace NaN / NaT with 0 in the last four columns
df.iloc[:, -4:] = df.iloc[:, -4:].fillna(0)

# ''' ****** Push to SQL Server ****** '''

# engine = db.create_engine('mssql+pyodbc://@TAXDB-PT001:1433/TestDB?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes')
# df.to_sql('MyGenTaxYR0045', engine,if_exists='replace')

# # SQL database connection string
# db_connection_string = 'mssql+pyodbc://TAXDB-PT001:1433/Budget_Levey_Data?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes'
# # Create SQLAlchemy engine
# engine = create_engine(db_connection_string)

# # Create a session
# Session = sessionmaker(bind=engine)
# session = Session()

# # Push the DataFrame to the SQL table
# df.to_sql('GenTaxYR0045', con=engine, if_exists='replace', index=False)

# print("Data pushed to SQL database successfully.")

# # Close the session
# session.close()

# Create a CSV with results from changes
df.to_csv(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2023 Example Reports\GenTaxYR045Modified.csv", index=False, header=True)
