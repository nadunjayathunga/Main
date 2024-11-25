import pandas as pd
from datetime import datetime,timedelta
import numpy as np
from dateutil.relativedelta import relativedelta
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt,RGBColor, Cm, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import calendar
import re
import random
import os

# PATH = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\data.csv'

# PATH_INV = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\fInvoices.csv'

# PATH_ROOM = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\room.csv'
# end_date :datetime = datetime(year=2024,month=10,day=31)
# start_date = end_date - relativedelta(months=12) + timedelta(days=1)
# df = pd.read_csv(PATH, parse_dates=['start_date', 'end_date'], date_parser=lambda x: datetime.strptime(x, '%m/%d/%Y'))
# df = df.loc[(df['end_date']>=start_date) & (df['start_date']<=end_date)]
# rooms: list = df['room_id'].unique()
# occupancy = {}
# for i in rooms:
#     periods = []
#     df_room = df.loc[(df['room_id']==i)]
#     for _, row in df_room.iterrows():
#         period: list = pd.date_range(start=row['start_date'] + pd.offsets.MonthEnd() ,end=row['end_date']+ pd.offsets.MonthEnd(), freq='ME').to_pydatetime().tolist()
#         period = [i for i in period if i >=start_date and i<= end_date]
#         periods += period
#     occupancy[i] = set(periods)

# cols = pd.date_range(start=start_date ,end=end_date, freq='ME').to_pydatetime().tolist()
# result_dict = {date: [False] * len(rooms) for date in cols}
# result_dict['room'] = rooms

# occupany_report: pd.DataFrame= pd.DataFrame(data=result_dict).set_index('room')
# timeperiods:list = list(occupany_report.columns)
# for room, row in occupany_report.iterrows():
#     for j in timeperiods:
#         if j in occupancy[room]:
#             occupany_report.loc[room,j] = True
# occupany_report.reset_index(inplace=True)
# print(occupany_report)

# pp_start:datetime = end_date - relativedelta(months=1)
# new_contracts:list = df.loc[(df['start_date'] >= datetime(year=end_date.year,month=end_date.month,day=1)) & (df['start_date']<=end_date),'order_id'].tolist()
# close_contracts:list = df.loc[(df['end_date'] >= datetime(year=pp_start.year,month=pp_start.month,day=1)) & (df['end_date']<datetime(year=end_date.year,month=end_date.month,day=1)),'order_id'].tolist()
# fInvoices = pd.read_csv(PATH_INV, parse_dates=['invoice_date'], date_parser=lambda x: datetime.strptime(x, '%m/%d/%Y'))

# new:pd.DataFrame = fInvoices.loc[(fInvoices['order_id'].isin(new_contracts)) & (fInvoices['invoice_date'] <= end_date) & (fInvoices['invoice_date']>=datetime(year=end_date.year,month=end_date.month,day=1)),['customer_code','amount']].groupby('customer_code',as_index=False)['amount'].sum()
# vacated:pd.DataFrame = fInvoices.loc[(fInvoices['order_id'].isin(close_contracts)) & (fInvoices['invoice_date'] <= pp_start) & (fInvoices['invoice_date']>=datetime(year=pp_start.year,month=pp_start.month,day=1)),['customer_code','amount']].groupby('customer_code',as_index=False)['amount'].sum()
# price_change:pd.DataFrame = fInvoices.loc[(fInvoices['invoice_date']>=datetime(year=pp_start.year,month=pp_start.month,day=1)) & (fInvoices['invoice_date']<=end_date),['customer_code','invoice_date','amount']]
# price_change = price_change.pivot_table(index='customer_code', columns=pd.Grouper(key='invoice_date',freq='ME'), values='amount').fillna(value=0).reset_index()

# price_change.loc[:,'change'] = price_change.iloc[:,2] - price_change.iloc[:,1]
# price_change = price_change.loc[price_change['change']!=0,['customer_code','change']]

end_date = datetime(year=2024,month=12,day=31)
end_date = end_date +  pd.offsets.MonthEnd(0)

print(end_date)

