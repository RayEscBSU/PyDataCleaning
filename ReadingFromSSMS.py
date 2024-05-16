from sqlalchemy import create_engine, MetaData, Table, select
from sqlalchemy.orm import sessionmaker

db_connection_string = 'mssql+pyodbc://@TAXDB-PT001:1433/Budget_Levey_Data?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes'
engine = create_engine(db_connection_string) #connects to db
connection = engine.connect()

# Reflect the existing table
metadata = MetaData(bind=engine)
county_category_detail = Table('CountyCategoryDetail', metadata, autoload=True, autoload_with=engine)

# Select all records from the table
query = select([county_category_detail])

# Execute the query and fetch the results
results = connection.execute(query).fetchall()

# Print results
for row in results:
    print(row)

# Close the connection
connection.close()