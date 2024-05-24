from sqlalchemy import create_engine, Column, Integer, String, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import pandas as pd

df = pd.read_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2024_County_Reports\Ada County (003) - Copy.xlsx")

# Trim leading and trailing whitespaces from all string values
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
# Trim leading and trailing whitespaces from the headers
df.columns = df.columns.str.strip()

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




#print
print(df.iloc[:10, :10])

#create a csv with results from changes
df.to_excel(r"C:\Users\rescobedo\OneDrive - State of Idaho\RayE\2024_County_Reports\Ada County (003) TestOutput.xlsx",sheet_name='Sheet1', index=False, header=True, engine='openpyxl')