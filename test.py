from sqlalchemy import create_engine, Column, Integer, String, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import pandas as pd

# Define existing SQL table model
Base = declarative_base()

class CountyCategoryDetail(Base):
    __tablename__ = 'CountyCategoryDetail'
    
    year = Column('Year', Integer, default=24)
    roll = Column('Roll', String, default='TEST')
    county_number = Column('County Number', Integer, default=1)
    county_name = Column('County Name', String, default='TEST')
    region_number = Column('Region Number', Integer, default=3)
    category = Column(' Category Number', Integer, default=101)
    acreage = Column(Float, default=0.0)  # Set a default value here
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
    nnet_value = Column('Net Value', Integer, default=0)
    residential_urban = Column('Residential Urban', Integer, default=0)
    residential_rural = Column('Residential Rural',Integer, default=0)
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

# Read Excel file
new_data_df = pd.read_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2023 Example Reports\Ada\Ada Bunderson Report (Unformatted)From Ada 2023.xlsx")

# Strip whitespace from DataFrame column names
new_data_df.columns = [col.strip() for col in new_data_df.columns]

# Define the column mapping: DataFrame column -> SQLAlchemy attribute
column_mapping = {
    'YEAR': 'year',
    'ROLL': 'roll',
    'COUNTY': 'county_number',
    'CATEGORY': 'category',
    'TOTAL_VALUE': 'market_value'
}

# Rename DataFrame columns based on mapping
new_data_df.rename(columns=column_mapping, inplace=True)

# Keep only the columns that have been mapped
new_data_df = new_data_df[list(column_mapping.values())]

# Convert DataFrame rows to SQLAlchemy objects
records = new_data_df.to_dict(orient='records')
county_details = []
for record in records:
    # Set defaults for missing mapped columns
    for mapped_col in column_mapping.values():
        if mapped_col not in record:
            column = getattr(CountyCategoryDetail, mapped_col, None)
            default_value = column.default.arg if column and column.default is not None else None
            record[mapped_col] = default_value

    county_detail = CountyCategoryDetail(**record)
    county_details.append(county_detail)

# Add new records to the database
session.add_all(county_details)
session.commit()
