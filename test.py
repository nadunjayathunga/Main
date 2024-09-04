import pandas as pd
from datetime import datetime
import numpy as np
# from dateutil.relativedelta import relativedelta

PATH_TS = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\fTimesheet.csv'
PATH_GL = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\merged.csv'
PATH_OT = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\fOT.csv'
PATH_EXCLUDE = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dExclude.csv'
PATH_EMP = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dEmployee.csv'

end_date :datetime = datetime(year=2024,month=7,day=31) 

fTimesheet:pd.DataFrame =  pd.read_csv(filepath_or_buffer=PATH_TS,dtype={'cost_center':str,'job_id':str},parse_dates=['v_date'],date_format='%m/%d/%Y')
fTimesheet = fTimesheet.loc[~fTimesheet['job_id'].isin(['discharged','not_joined'])]
fTimesheet.loc[:,'v_date'] = fTimesheet['v_date'] + pd.offsets.MonthEnd(0)

dEmployee:pd.DataFrame =  pd.read_csv(filepath_or_buffer=PATH_EMP,usecols=['Employee_Code'],index_col='Employee_Code')

fGL:pd.DataFrame = pd.read_csv(filepath_or_buffer=PATH_GL,parse_dates=['Voucher Date'],date_format='%m/%d/%Y')
fGL.loc[:,'Voucher Date'] = fGL['Voucher Date'] + pd.offsets.MonthEnd(0)


dExclude :pd.DataFrame = pd.read_csv(PATH_EXCLUDE,usecols=['dc_emp_beni','dc_trpt','dc_out','dc_sal','job_type'])




cogs_ledger_mapping = {'dc_emp_beni':[5010105002,5010101006,5010101005,5010105001,5010105003,5010101007,5010103002,5010103001,5010101004],
                       'dc_trpt':[5010102001,5010102002],
                       'dc_out':[5010101008],'dc_sal':[5010101001]}

def job_profitability(fTimesheet:pd.DataFrame,fGL:pd.DataFrame,end_date:datetime,dEmployee:pd.DataFrame,dExclude:pd.DataFrame)->pd.DataFrame:
    start_date:datetime = datetime(year=end_date.year,month=1,day=1)
    dEmployee.reset_index(inplace=True)
    fGL = fGL.loc[:,['Cost Center','Voucher Date','Ledger_Code','Amount','Third_Level_Group_Name','Second_Level_Group_Name']]
    fGL = fGL.loc[(fGL['Voucher Date']>=start_date) & (fGL['Voucher Date']<=end_date) & 
                  (fGL['Second_Level_Group_Name'] == 'Manpower Cost') & 
                  (~fGL['Ledger_Code'].isin([5010101002,5010101003])),['Cost Center','Voucher Date','Ledger_Code','Amount']]
    emp_list :list = dEmployee['Employee_Code'].tolist()
    fGL_emp :pd.DataFrame = fGL.loc[fGL['Cost Center'].isin(emp_list)]
    fGL_emp = fGL_emp.groupby(by=['Cost Center','Voucher Date','Ledger_Code'],as_index=False)['Amount'].sum()
    fGL_emp = fGL_emp.loc[fGL_emp['Amount']!=0]
    # print('fGL_emp')
    # print(fGL_emp)
    # fGL.fillna(value='Dummy',inplace=True)
    # fGL = fGL.groupby(by=['Voucher Date','Cost Center','Ledger_Code'])['Amount'].sum().unstack(fill_value=0).reset_index([0,1])
    # fGL_emp = fGL_emp.loc[:,(fGL_emp!=0).any(axis=0)]
    # fGL_oh :pd.DataFrame = fGL.loc[~fGL['Cost Center'].isin(emp_list)]
    # fGL_oh = fGL_oh.loc[:,(fGL_oh!=0).any(axis=0)]
    fTimesheet = fTimesheet.groupby(['cost_center', 'job_id', 'v_date']).size().reset_index(name='count')

    timesheet_sum :dict = {'dc_emp_beni':None,'dc_trpt':None,'dc_out':None,'dc_sal':None}
    timesheet_jobs :dict = {'dc_emp_beni':None,'dc_trpt':None,'dc_out':None,'dc_sal':None}
    billable_jobs:list = fTimesheet.loc[fTimesheet['job_id'].str.contains('ESS/CTR'),'job_id'].unique().tolist()
    
    for c in dExclude.columns:
        if c != 'job_type':
            valid_jobs = dExclude.loc[dExclude[c]==False]['job_type'].tolist() + billable_jobs
            timesheet_sum[c]  = fTimesheet.loc[fTimesheet['job_id'].isin(valid_jobs)].groupby(['cost_center','v_date'],as_index=False)['count'].sum()
            # print(timesheet_sum[c])
            timesheet_jobs[c] = fTimesheet.loc[fTimesheet['job_id'].isin(valid_jobs)]

    allocation_dict :dict = {}
    for _,i in fGL_emp.iterrows():
        df_type :str = [(k,v) for k,v in cogs_ledger_mapping.items() if i['Ledger_Code'] in v][0][0]
        print(df_type)
        df :pd.DataFrame = timesheet_sum[df_type]
        # print('df')
        print(df)
        # try:
        total_days: int = df.loc[(df['v_date'] == i['Voucher Date']) & (df['cost_center'] == i['Cost Center']),'count'].iloc[0]
        # print(total_days)
        # except:
            # allocation_dict:dict ={ i['Voucher Date']:{'unallocated':i['Amount']}}
            # print(allocation_dict)
        timesheet_detailed:pd.DataFrame = timesheet_jobs[df_type]
        # print('timeshet_detailed')
        print(timesheet_detailed)
        timesheet_detailed = timesheet_detailed.loc[(timesheet_detailed['cost_center'] == i['Cost Center']) & (timesheet_detailed['v_date']==i['Voucher Date']),['job_id','count']]
        allocation_dict_init = {}
        for _,j in timesheet_detailed.iterrows():
            allocated :float =i['Amount']/total_days * j['count']
            print(f"{i['Amount']}:{total_days}:{j['count']}")
            allocation_dict_init[j['job_id']] =  allocated
            # print(allocation_dict_init[j['job_id']])
            print(allocation_dict_init)
        allocation_dict[i['Voucher Date']] = allocation_dict_init
    print(allocation_dict)

job_profitability(fTimesheet=fTimesheet,fGL=fGL,end_date=end_date,dEmployee=dEmployee,dExclude=dExclude)

# TODO NEED A DATAFRAME FOR MONTHWISE / EMPLOYEE WISE TOTAL NUMBER OF DAYS WORKED

