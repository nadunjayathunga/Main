from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
from data import company_data,company_info,db_info,table_info
from sqlalchemy import create_engine,text
import psycopg2

PATH = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\df_rating.xlsx'
df = pd.read_excel(io=PATH,sheet_name='df_rating_3')
df1 = pd.read_excel(io=PATH,sheet_name='df_rating_0')

first_row = df.head(1).iloc[0]
customer = df.head(1)['Customer Name'].iloc[0]
x = first_row[first_row == 'Yes'].index[0] if any(first_row == 'Yes') else None
start: int = df1.columns.tolist().index(x)
name = df1.loc[df1['Customer Name']==customer].index[0]

mystr: str = ''
for i,j in enumerate(range(start,start+3)):
    value = df1.iloc[name ,j]
    date = df1.columns.tolist()[j].date()
    mystr += f'{i+1}: {date}-{value:,.0f}{"/ "  if i<2 else ""}'
print(mystr)

