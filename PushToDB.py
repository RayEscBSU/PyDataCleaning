import pandas as pd
from sqlalchemy import create_engine

# Create the connection string
db_connection_string = 'mssql+pyodbc://TAXDB-PT001/Budget_Levey_Data?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes'

# Create an SQLAlchemy engine
engine = create_engine(db_connection_string)

# Path to your CSV file
csv_file_path = r'C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2023 Example Reports\GenTaxYR045Modified.csv'

# Read the CSV file into a DataFrame
df = pd.read_csv(csv_file_path)

# Push the DataFrame to SQL Server
df.to_sql('MyGenTaxYR045', con=engine, if_exists='replace', index=False)

print("CSV data has been pushed to SQL Server.")
