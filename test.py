import pandas as pd
from datetime import datetime
import numpy as np
from dateutil.relativedelta import relativedelta

PATH_TS = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\fTimesheet.csv'
PATH_GL = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\merged.csv'
PATH_OT = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\fOT.csv'
PATH_EXCLUDE = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dExclude.csv'
PATH_EMP = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dEmployee.csv'
PATH_INV = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\fInvoices.csv'
PATH_JOB = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dJobs.csv'
PATH_CUS = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\NBNL\dCustomers.csv'


end_date :datetime = datetime(year=2024,month=7,day=31) 

fTimesheet:pd.DataFrame =  pd.read_csv(filepath_or_buffer=PATH_TS,dtype={'cost_center':str,'job_id':str},parse_dates=['v_date'],date_format='%m/%d/%Y')
fTimesheet['v_date'] = pd.to_datetime(fTimesheet['v_date'])

fTimesheet = fTimesheet.loc[~fTimesheet['job_id'].isin(['discharged','not_joined'])]
fTimesheet.loc[:,'v_date'] = fTimesheet['v_date'] + pd.offsets.MonthEnd(0)

dEmployee:pd.DataFrame =  pd.read_csv(filepath_or_buffer=PATH_EMP,usecols=['Employee_Code'],index_col='Employee_Code')

fGL:pd.DataFrame = pd.read_csv(filepath_or_buffer=PATH_GL,parse_dates=['Voucher Date'],date_format='%m/%d/%Y')
fGL['Voucher Date'] = pd.to_datetime(fGL['Voucher Date'])
fGL.loc[:,'Voucher Date'] = fGL['Voucher Date'] + pd.offsets.MonthEnd(0)

dExclude :pd.DataFrame = pd.read_csv(PATH_EXCLUDE,usecols=['dc_emp_beni','dc_trpt','dc_out','dc_sal','job_type'])

fOT = pd.read_csv(PATH_OT)
fOT['date'] = pd.to_datetime(fOT['date'])
fInvoices :pd.DataFrame = pd.read_csv(PATH_INV)
fInvoices['Invoice_Date'] = pd.to_datetime(fInvoices['Invoice_Date'])
dJobs = pd.read_csv(PATH_JOB)
dCustomers:pd.DataFrame =pd.read_csv(PATH_CUS)

cogs_ledger_mapping = {'dc_emp_beni':[5010105002,5010101006,5010101005,5010105001,5010105003,5010101007,5010103002,5010103001,5010101004],
                       'dc_trpt':[5010102001,5010102002],
                       'dc_out':[5010101008],'dc_sal':[5010101001]}

def job_profitability(fTimesheet:pd.DataFrame,fGL:pd.DataFrame,end_date:datetime,dEmployee:pd.DataFrame,dExclude:pd.DataFrame,fOT:pd.DataFrame,fInvoices:pd.DataFrame,cogs_map:dict,dJobs:pd.DataFrame)->pd.DataFrame:

    start_date:datetime = datetime(year=end_date.year,month=1,day=1)
    periods :list =  pd.date_range(start=start_date, end=end_date, freq='ME').to_pydatetime().tolist()
    fGL = fGL.loc[:,['Cost Center','Voucher Date','Ledger_Code','Amount','Third_Level_Group_Name','Second_Level_Group_Name']]
    fGL = fGL.loc[~fGL['Ledger_Code'].isin([5010101002,5010101003])]
    emp_list :list = dEmployee.index.tolist()
    timesheet_sum :dict = {'dc_emp_beni':None,'dc_trpt':None,'dc_out':None,'dc_sal':None}
    timesheet_jobs :dict = {'dc_emp_beni':None,'dc_trpt':None,'dc_out':None,'dc_sal':None}
    periodic_allocation :dict = {}

    for period in periods:
        st_date :datetime = period + relativedelta(day=1)
        fGL_fitlered :pd.DataFrame = fGL.loc[(fGL['Voucher Date']>=st_date) & (fGL['Voucher Date']<=period) & 
                    (fGL['Second_Level_Group_Name'] == 'Manpower Cost') ,['Cost Center','Voucher Date','Ledger_Code','Amount']]
        fGL_emp :pd.DataFrame = fGL_fitlered.loc[fGL_fitlered['Cost Center'].isin(emp_list)]
        fGL_emp = fGL_emp.groupby(by=['Cost Center','Voucher Date','Ledger_Code'],as_index=False)['Amount'].sum()
        fGL_emp = fGL_emp.loc[fGL_emp['Amount']!=0]
        fTimesheet_filtered = fTimesheet.loc[(fTimesheet['v_date'] >= st_date) & (fTimesheet['v_date']<=period)]
        fTimesheet_filtered = fTimesheet_filtered.groupby(['cost_center', 'job_id', 'v_date']).size().reset_index(name='count')
        billable_jobs:list = fTimesheet_filtered.loc[fTimesheet_filtered['job_id'].str.contains('ESS/CTR'),'job_id'].unique().tolist()
        
        for c in dExclude.columns:
            if c != 'job_type':
                valid_jobs = dExclude.loc[dExclude[c]==False]['job_type'].tolist() + billable_jobs
                timesheet_sum[c]  = fTimesheet_filtered.loc[fTimesheet_filtered['job_id'].isin(valid_jobs)].groupby(['cost_center','v_date'],as_index=False)['count'].sum()
                timesheet_jobs[c] = fTimesheet_filtered.loc[fTimesheet_filtered['job_id'].isin(valid_jobs)]
        allocation_dict :dict = {}
        unallocated_amount :float = 0
        for _,i in fGL_emp.iterrows():
            df_type :str = [(k,v) for k,v in cogs_map.items() if i['Ledger_Code'] in v][0][0]
            df_sum :pd.DataFrame = timesheet_sum[df_type]
            timesheet_detailed:pd.DataFrame = timesheet_jobs[df_type]
            try:
                total_days: int = df_sum.loc[(df_sum['v_date'] == i['Voucher Date']) & (df_sum['cost_center'] == i['Cost Center']),'count'].iloc[0]
                timesheet_detailed = timesheet_detailed.loc[(timesheet_detailed['v_date']==i['Voucher Date']) & (timesheet_detailed['cost_center'] == i['Cost Center']),['job_id','count']]
                allocation_dict_init = {}
                for _,j in timesheet_detailed.iterrows():
                    allocated :float =i['Amount']/total_days * j['count']
                    allocation_dict_init[j['job_id']] =  allocated
                allocation_dict = {k: allocation_dict_init.get(k,0) + allocation_dict.get(k,0) for k in set(allocation_dict)|set(allocation_dict_init)}
            except IndexError:
                unallocated_amount += i['Amount']
                allocation_dict['Un-Allocated'] = unallocated_amount
        fOT_filtered :pd.DataFrame = fOT.loc[(fOT['date'] >= st_date) & (fOT['date']<=period)]
        fOT_filtered :dict= fOT_filtered.groupby(by='job_id')['net'].sum().to_dict()
        allocation_dict = {k:allocation_dict.get(k,0) + fOT.get(k,0) for k in set(allocation_dict)|set(fOT_filtered)}
        inv_filtered_cust :dict= fInvoices.loc[(fInvoices['Invoice_Date'] >= st_date) & (fInvoices['Invoice_Date']<=period),['Order_ID','Net_Amount']].groupby('Order_ID')['Net_Amount'].sum().to_dict()
        allocation_dict = {k:allocation_dict.get(k,0) + inv_filtered_cust.get(k,0) for k in set(allocation_dict)|set(inv_filtered_cust)}
        periodic_allocation[period] = allocation_dict
    cy_cp:pd.DataFrame = pd.DataFrame(list(periodic_allocation[end_date].items()),columns=['Order_ID','Amount'])
    cy_cp = pd.merge(left=cy_cp,right=dJobs[['Order_ID','Customer_Code','Employee_Code']],on='Order_ID',how='left')
    cy_cp_cus :pd.DataFrame = cy_cp.groupby(by='Customer_Code',as_index=False)['Amount'].sum()
    cy_cp_emp :pd.DataFrame= cy_cp.groupby(by='Employee_Code',as_index=False)['Amount'].sum()
    cy_ytd:pd.DataFrame = pd.DataFrame()
    for period in periods:
        month_df :pd.DataFrame = pd.DataFrame(list(periodic_allocation[period].items()),columns=['Order_ID','Amount'])
        cy_ytd = pd.concat([month_df,cy_ytd])
    cy_ytd = pd.merge(left=cy_ytd,right=dJobs[['Order_ID','Customer_Code','Employee_Code']],on='Order_ID',how='left')
    cy_ytd_cus:pd.DataFrame = cy_ytd.groupby(by='Customer_Code',as_index=False)['Amount'].sum()
    cy_ytd_emp:pd.DataFrame = cy_ytd.groupby(by='Employee_Code',as_index=False)['Amount'].sum()
    return {'periodic_allocation':periodic_allocation,'cy_cp_cus':cy_cp_cus,'cy_ytd_cus':cy_ytd_cus,'cy_cp_emp':cy_cp_emp,'cy_ytd_emp':cy_ytd_emp}

profitability = job_profitability(fTimesheet=fTimesheet,fGL=fGL,end_date=end_date,dEmployee=dEmployee,dExclude=dExclude,fOT=fOT,fInvoices=fInvoices,cogs_map=cogs_ledger_mapping,dJobs=dJobs)
print(profitability['cy_cp_emp'])
print(profitability['cy_ytd_emp'])

# last_month :pd.DataFrame = pd.DataFrame(list(x[end_date].items()),columns=['Order_ID','Amount']) # for end_date period

# last_month = pd.merge(left=last_month,right=dJobs[['Order_ID','Customer_Code']],on='Order_ID',how='left')

# last_month = last_month.groupby(by='Customer_Code')['Amount'].sum()

# start_date:datetime = datetime(year=end_date.year,month=1,day=1)
# periods :list =  pd.date_range(start=start_date, end=end_date, freq='ME').to_pydatetime().tolist()
# ytd_df = pd.DataFrame()
# for period in periods:
#     month_df :pd.DataFrame = pd.DataFrame(list(x[end_date].items()),columns=['Order_ID','Amount'])
#     ytd_df = pd.concat([month_df,ytd_df])

# ytd_df = pd.merge(left=ytd_df,right=dJobs[['Order_ID','Customer_Code']],on='Order_ID',how='left')

# ytd_df = ytd_df.groupby(by='Customer_Code')['Amount'].sum()

# x_inter= [list(i.items()) for i in x.values()]
# y = [j for i in x_inter for j in i]
# df = pd.DataFrame(y)
# print(df)

# customer_list: list = sorted(fInvoices.loc[(fInvoices['Invoice_Date'] >= datetime(year=end_date.year,
#                                                                                   month=end_date.month, day=1)) & (
#                                                    fInvoices['Invoice_Date'] <= end_date), 'Cus_Name'].unique())
# cy_cp:pd.DataFrame = profitability['cy_cp']
# for i in customer_list:
#     id = dCustomers.loc[(dCustomers['Cus_Name']==i),'Customer_Code'].tolist()
#     profit_cp :float = cy_cp.loc[cy_cp['Customer_Code'].isin(id),'Amount'].sum()
#     print(profit_cp)
    