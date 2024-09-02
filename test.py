import pandas as pd
from datetime import datetime
# from dateutil.relativedelta import relativedelta

PATH_TS = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\fTimesheet.csv'
PATH_GL = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\merged.csv'
PATH_OT = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\fOT.csv'
PATH_EXCLUDE = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dExclude.csv'
PATH_EMP = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dEmployee.csv'
end_date :datetime = datetime(year=2024,month=7,day=31) 

fTimesheet:pd.DataFrame =  pd.read_csv(filepath_or_buffer=PATH_TS,dtype={'cost_center':str,'job_id':str},parse_dates=['v_date'],date_format='%m/%d/%Y')
fTimesheet = fTimesheet.loc[~fTimesheet['job_id'].isin(['discharged','not_joined'])]
fTimesheet.loc[:,'v_date'] = fTimesheet['v_date'] + pd.offsets.MonthEnd()

dEmployee:pd.DataFrame =  pd.read_csv(filepath_or_buffer=PATH_EMP,usecols=['Employee_Code'],index_col='Employee_Code')

fGL:pd.DataFrame = pd.read_csv(filepath_or_buffer=PATH_GL,parse_dates=['Voucher Date'],date_format='%m/%d/%Y')
fGL.loc[:,'Voucher Date'] = fGL['Voucher Date'] + pd.offsets.MonthEnd()

def job_profitability(fTimesheet:pd.DataFrame,fGL:pd.DataFrame,end_date:datetime,dEmployee:pd.DataFrame)->pd.DataFrame:
    start_date:datetime = datetime(year=end_date.year,month=1,day=1)
    dEmployee.reset_index(inplace=True)
    fGL = fGL.loc[:,['Cost Center','Voucher Date','Ledger_Code','Amount','Third_Level_Group_Name','Second_Level_Group_Name']]
    fGL = fGL.loc[(fGL['Voucher Date']>=start_date) & (fGL['Voucher Date']<=end_date) & (fGL['Second_Level_Group_Name'] == 'Manpower Cost') &(~fGL['Ledger_Code'].isin(5010101002,5010101003)),['Cost Center','Voucher Date','Ledger_Code','Amount']]
    fGL.fillna(value='Dummy',inplace=True)
    fGL_emp = fGL.groupby(by=['Cost Center','Voucher Date','Ledger_Code'])['Amount'].sum()
    fGL_emp = fGL_emp.loc[fGL_emp !=0]

    fGL = fGL.groupby(by=['Voucher Date','Cost Center','Ledger_Code'])['Amount'].sum().unstack(fill_value=0).reset_index([0,1])
    emp_list :list = dEmployee['Employee_Code'].tolist()
    # fGL_emp :pd.DataFrame = fGL.loc[fGL['Cost Center'].isin(emp_list)]
    # fGL_emp = fGL_emp.loc[:,(fGL_emp!=0).any(axis=0)]
    fGL_oh :pd.DataFrame = fGL.loc[~fGL['Cost Center'].isin(emp_list)]
    fGL_oh = fGL_oh.loc[:,(fGL_oh!=0).any(axis=0)]
    fTimesheet = fTimesheet.groupby(['cost_center', 'job_id', 'v_date']).size().reset_index(name='count')
    fTimesheet.to_csv('fTimesheet.csv')


job_profitability(fTimesheet=fTimesheet,fGL=fGL,end_date=end_date,dEmployee=dEmployee)

# TODO NEED A DATAFRAME FOR MONTHWISE / EMPLOYEE WISE TOTAL NUMBER OF DAYS WORKED
