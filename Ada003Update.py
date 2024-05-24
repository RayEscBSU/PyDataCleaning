from sqlalchemy import create_engine, Column, Integer, String, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import pandas as pd

# Define existing SQL table model
Base = declarative_base()

class CountyCategoryDetail(Base):
    __tablename__ = 'CountyCategoryDetail'
     
    year = Column('Year', Integer, default=2024)
    roll = Column('Roll', String, default='TEST2')
    county_number = Column('County Number', Integer, default=1)
    county_name = Column('County Name', String, default='test2')
    region_number = Column('Region Number', Integer, default=3)
    category = Column(' Category Number', Integer, default=101)
    acreage = Column(Float, default=0.0)
    market_value = Column('Market Value', Integer, default=0)
    homeowner_s_exemption = Column("Homeowner's Exemption", Integer, default=0)
    increment = Column('Increment', Integer, default=0)
    personal_property_exemption = Column('Personal Property Exemption', Integer, default=0)
    hardship = Column('Hardship', Integer, default=0)
    pollution = Column('Pollution', Integer, default=0)
    new_capital_investment = Column('New Capital Investment (63-4502)', Integer, default=0)
    business_investment = Column ('Business Investment (63-602NN)', Integer, default=0)
    site_improvement = Column( 'Site Improvement', Integer, default=0)
    casualty_loss = Column('Casualty Loss (63-602X)', Integer, default=0)
    qie = Column('QIE', Integer, default=0)
    net_value = Column('Net Value', Integer, default=0)
    residential_urban = Column('Residential Urban', Integer, default=0)
    residential_rural = Column('Residential Rural', Integer, default=0)
    commercial_urban = Column('Commercial Urban', Integer, default=0)
    commercial_rural = Column('Commercial Rural', Integer, default=0)
    ag = Column('Ag', Integer, default=0)
    timber = Column('Timber', Integer, default=0)
    mining = Column('Mining', Integer, default=0)
    id = Column(Integer, primary_key=True)

# Create session and engine - Database connection details
db_connection_string = 'mssql+pyodbc://@TAXDB-PT001:1433/Budget_Levey_Data?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes'
engine = create_engine(db_connection_string)
Session = sessionmaker(bind=engine)
session = Session()

#read excel file
df = pd.read_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2024_County_Reports\Ada County (003) - Copy.xlsx")

# Trim leading and trailing whitespaces from all string values
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
# Trim leading and trailing whitespaces from the headers
df.columns = df.columns.str.strip()

"""
    The column mapping will specify which columns from the CSV I want to read
    all others will be ignored.  
"""
# Define the column mapping: DataFrame column -> SQLAlchemy attribute
column_mapping = {
    'Category': 'category',
    'Acres': 'acreage',
    'Market Value': 'market_value',
    'Homeowners' : 'homeowner_s_exemption',
    'Increment' : 'increment',
    'Hardship' : 'hardship',
    'QIE' : 'qie',
    'Pollution' : 'pollution',
    'New Capital Investment' : 'new_cap_investment', 
    'Casualty' : 'casualty_loss',
    'Site Improvments' : 'site_improvement',
    'PP Exemption' : 'personal_property_exemption',
    'Taxable Value' : 'net_value'
}

# Rename DataFrame columns based on mapping and keep only those columns
df = df.rename(columns=column_mapping).loc[:, column_mapping.values()]

"""
    Columns that exist in SQl database but not found in column mapping
    will be updated with the defualt value set above   
 """
# get columns and default values from sqlalchemy table
for column, default_value in CountyCategoryDetail.__table__.columns.items():
    # check if columns is mapped and but not found in csv 
    if column in column_mapping.values() and column not in df:
        # set default value it one is assigned, if no default is assigned set to none
        df[column] = default_value.default.arg if default_value.default is not None else None 

# Convert DataFrame rows to SQLAlchemy objects
county_details = [CountyCategoryDetail(**record) for record in df.to_dict(orient='records')]

# Add new records to the database
session.add_all(county_details)
session.commit()
session.close()



#print
#print(df.iloc[:10, :10])

#create a csv with results from changes
#df.to_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2024_County_Reports\Ada County (003) TestOutput.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')