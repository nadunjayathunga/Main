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

fTimesheet = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\ftimesheettest.csv'
fTimesheet = pd.read_csv(fTimesheet)
fTimesheet['v_date'] = pd.to_datetime(fTimesheet['v_date'])
fGL = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\fgltest.csv'
fGL = pd.read_csv(fGL)
fGL['voucher_date'] = pd.to_datetime(fGL['voucher_date'])
dEmployee = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\demployeetest.csv'
dEmployee = pd.read_csv(dEmployee)
dEmployee['dob'] = pd.to_datetime(dEmployee['dob'])
dEmployee['doj'] = pd.to_datetime(dEmployee['doj'])
dEmployee['termination_date'] = pd.to_datetime(dEmployee['termination_date'])
dExclude = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\dexcludetest.csv'
dExclude = pd.read_csv(dExclude)
fOT = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\fottest.csv'
fOT = pd.read_csv(fOT)
fOT['voucher_date'] = pd.to_datetime(fOT['voucher_date'])
dJobs = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\djobstest.csv'
dJobs = pd.read_csv(dJobs)
fMI = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\fmitest.csv'
fMI = pd.read_csv(fMI)
fInvoices = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\finvoicestest.csv'
fInvoices = pd.read_csv(fInvoices)
fInvoices['invoice_date'] = pd.to_datetime(fInvoices['invoice_date'])


end_date = datetime(year=2024,month=10,day=31)
cogs_ledger_map :dict = {'elite_security':{'dc_emp_beni':[5010105002,5010101006,5010101005,5010105001,5010105003,5010101007,5010103002,5010103001,5010101004],
                       'dc_trpt':[5010102001,5010102002],
                       'dc_out':[5010101008],'dc_sal':[5010101001]},
                       'premium':{'dc_emp_beni':[5010101008,5010101009,5010101010,5010103002,5010104001,5010105002,5010105006],
                       'dc_trpt':[5010102001,5010102002],
                       'dc_out':[5010101011],'dc_sal':[5010101005]}}

def job_profitability(fTimesheet: pd.DataFrame, fGL: pd.DataFrame, end_date: datetime, dEmployee: pd.DataFrame,
                      dExclude: pd.DataFrame, fOT: pd.DataFrame, fInvoices: pd.DataFrame, cogs_map: dict,
                      dJobs: pd.DataFrame, database, fMI: pd.DataFrame) -> dict:
    start_date: datetime = datetime(year=end_date.year, month=10, day=1)
    periods: list = pd.date_range(start=start_date, end=end_date, freq='ME').to_pydatetime().tolist()
    dEmployee.set_index(keys='emp_id',inplace=True)
    fGL = fGL.loc[:,
          ['cost_center', 'voucher_date', 'ledger_code', 'amount', 'third_level', 'second_level', 'first_level']]
    cy_cp_cus_np = None
    cy_cp_emp_np = None
    cy_ytd_emp_np = None
    cy_ytd_cus_np = None
    nbnl_profitability = None

    if database in ['elite_security', 'premium']:
        if database == 'elite_security':
            # excluding Direct Cost - Normal OT and Direct Cost - Holiday OT as balances in those ledgers are treated separately.
            fGL = fGL.loc[~fGL['ledger_code'].isin([5010101002, 5010101003])]
        else:
            fGL = fGL.loc[~fGL['ledger_code'].isin([5010101006, 5010101007])]
        emp_list_full: list = dEmployee.index.tolist()
        driversandcleaners: list = dEmployee.loc[
            dEmployee['designation'].isin(['HEAVY DRIVER', 'DRIVER', 'CAMP SUPERVISOR'])].index.tolist()
        emp_list: list = [i for i in emp_list_full if i not in driversandcleaners]
        timesheet_sum: dict = {'dc_emp_beni': None, 'dc_trpt': None, 'dc_out': None, 'dc_sal': None}
        timesheet_jobs: dict = {'dc_emp_beni': None, 'dc_trpt': None, 'dc_out': None, 'dc_sal': None}
        timesheet_grand_sum: dict = {'dc_emp_beni': None, 'dc_trpt': None, 'dc_out': None, 'dc_sal': None}
        periodic_allocation: dict = {}

        for period in periods:
            consumable: dict = \
            fMI.loc[fMI['voucher_date'] == period, ['order_id', 'amount']].groupby(by='order_id', as_index=True)[
                'amount'].sum().to_dict()

            st_date: datetime = period + relativedelta(day=1)
            fGL_fitlered: pd.DataFrame = fGL.loc[(fGL['voucher_date'] >= st_date) & (fGL['voucher_date'] <= period) &
                                                 (fGL['second_level'] == 'Manpower Cost'), ['cost_center',
                                                                                            'voucher_date',
                                                                                            'ledger_code',
                                                                                            'amount']]
            fGL_emp: pd.DataFrame = fGL_fitlered.loc[fGL_fitlered['cost_center'].isin(emp_list)].groupby(
                by=['cost_center', 'voucher_date', 'ledger_code'], as_index=False)['amount'].sum()
            fGL_other: pd.DataFrame = \
                fGL_fitlered.loc[~fGL_fitlered['cost_center'].isin(emp_list), ['amount', 'ledger_code']].groupby(
                    'ledger_code',
                    as_index=False)[
                    'amount'].sum()
            fGL_other.to_csv('fGL_other.csv')
            fGL_emp = fGL_emp.loc[fGL_emp['amount'] != 0]
            fGL_emp.to_csv('fGL_emp.csv')
            # TODO You may group this to cogs map using the ledger code. to be fixed. it will reduce the no of iteretion by approx 12.5%
            fTimesheet_filtered: pd.DataFrame = fTimesheet.loc[
                (fTimesheet['v_date'] >= st_date) & (fTimesheet['v_date'] <= period)]
            # count of each combination 
            #     cost_center      order_id     v_date  count
            # 0       PH00001  Annual Leave 2024-01-31     28
            # 1       PH00001  Unpaid Leave 2024-01-31      1
            # 2       PH00001     WK-Worked 2024-01-31      2
            # 3       PH00002        OF-Off 2024-01-31      4
            # 4       PH00002  PH/CTR220002 2024-01-31     27
            fTimesheet_filtered = fTimesheet_filtered.groupby(['cost_center', 'order_id', 'v_date']).size().reset_index(
                name='count')
            billable_jobs: list = fTimesheet_filtered.loc[
                fTimesheet_filtered['order_id'].str.contains('ESS/CTR|PH/CTR'), 'order_id'].unique().tolist()
            for c in dExclude.columns:
                if c not in ['job_type', 'group']:
                    valid_jobs: list = dExclude.loc[dExclude[c] == False]['job_type'].tolist() + billable_jobs
                    timesheet_sum[c] = \
                        fTimesheet_filtered.loc[fTimesheet_filtered['order_id'].isin(valid_jobs)].groupby(
                            ['cost_center', 'v_date'], as_index=False)['count'].sum()
                    #     cost_center     v_date  count
                    # 0       PH00001 2024-01-31     31
                    # 1       PH00002 2024-01-31     31
                    # 2       PH00003 2024-01-31     31
                    # 3       PH00004 2024-01-31     31
                    # 4       PH00007 2024-01-31     31
                    timesheet_jobs[c] = fTimesheet_filtered.loc[fTimesheet_filtered['order_id'].isin(valid_jobs)]
                    #     cost_center      order_id     v_date  count
                    # 0       PH00001  Annual Leave 2024-01-31     28
                    # 1       PH00001  Unpaid Leave 2024-01-31      1
                    # 2       PH00001     WK-Worked 2024-01-31      2
                    # 3       PH00002        OF-Off 2024-01-31      4
                    # 4       PH00002  PH/CTR220002 2024-01-31     27
                    timesheet_grand_sum[c] = timesheet_sum[c]['count'].sum()

            allocation_dict: dict = {}
            allocation_dict = allocation_dict | consumable
            unallocated_amount: float = 0
            for _, i in fGL_emp.iterrows():
                df_type: str = next((ledger_type for ledger_type, values in cogs_ledger_map[database].items() if
                                     i['ledger_code'] in values))
                # TODO (a) YOU MAY FILTER df_sum/timesheet_sum and timesheet_detailed/timesheet_jobs only for those cost_centers apperiring in fGL_Emp. which will reduce the number of iterations.
                # Also filter by the ledger as well 
                df_sum: pd.DataFrame = timesheet_sum[df_type]
                timesheet_detailed: pd.DataFrame = timesheet_jobs[df_type]
                try:
                    total_days: int = df_sum.loc[(df_sum['v_date'] == i['voucher_date']) & (
                            df_sum['cost_center'] == i['cost_center']), 'count'].iloc[0]
                    timesheet_detailed = timesheet_detailed.loc[(timesheet_detailed['v_date'] == i['voucher_date']) & (
                            timesheet_detailed['cost_center'] == i['cost_center']), ['order_id', 'count']]
                    allocation_dict_init = {}
                    for _, j in timesheet_detailed.iterrows():
                        # TODO (a) only those cost centers having a value will return a value from below. 
                        allocated: float = i['amount'] / total_days * j['count']
                        allocation_dict_init[j['order_id']] = allocated
                    allocation_dict = {k: allocation_dict_init.get(k, 0) + allocation_dict.get(k, 0) for k in
                                       set(allocation_dict) | set(allocation_dict_init)}

                except IndexError:
                    unallocated_amount += i['amount']
                    allocation_dict['Un-Allocated'] = unallocated_amount
            fOT_filtered: pd.DataFrame = fOT.loc[(fOT['voucher_date'] >= st_date) & (fOT['voucher_date'] <= period)]
            fOT_filtered: dict = fOT_filtered.groupby(by='order_id')['amount'].sum().to_dict()
            allocation_dict = {k: allocation_dict.get(k, 0) + fOT_filtered.get(k, 0) for k in
                               set(allocation_dict) | set(fOT_filtered)}
            inv_filtered_cust: dict = fInvoices.loc[
                (fInvoices['invoice_date'] >= st_date) & (fInvoices['invoice_date'] <= period), ['order_id',
                                                                                                 'amount']].groupby(
                'order_id')['amount'].sum().to_dict()
            allocation_dict = {k: allocation_dict.get(k, 0) + inv_filtered_cust.get(k, 0) for k in
                               set(allocation_dict) | set(inv_filtered_cust)}
            for i in cogs_map[database]:
                z: float = fGL_other.loc[fGL_other['ledger_code'].isin(cogs_map[database][i])]['amount'].sum()
                if z != 0:
                    for _, row in timesheet_jobs[i].groupby(by='order_id', as_index=False)['count'].sum().iterrows():
                        overhead_allocation: dict = {}
                        value: float = z / timesheet_grand_sum[i] * row['count']
                        overhead_allocation[row['order_id']] = value
                        allocation_dict = {k: allocation_dict.get(k, 0) + overhead_allocation.get(k, 0) for k in
                                           set(allocation_dict) | set(overhead_allocation)}

            acc_types: list = dExclude.loc[dExclude['group'].isin(['Accommodation']), 'job_type'].tolist()
            accommodation_cost: float = sum([v for k, v in allocation_dict.items() if k in acc_types])
            non_accomo_sum: int = fTimesheet_filtered.loc[~fTimesheet_filtered['order_id'].isin(acc_types)][
                'count'].sum()
            non_accomo: pd.DataFrame = fTimesheet_filtered.loc[~fTimesheet_filtered['order_id'].isin(acc_types)]
            for _, row in non_accomo.iterrows():
                accommodation_allocation: dict = {}
                value: float = accommodation_cost / non_accomo_sum * row['count']
                accommodation_allocation[row['order_id']] = value
                allocation_dict = {k: allocation_dict.get(k, 0) + accommodation_allocation.get(k, 0) for k in
                                   set(allocation_dict) | set(accommodation_allocation)}


            if 'AC-ACCOMODATION' in allocation_dict:
                del allocation_dict['AC-ACCOMODATION']

            if 'AC' in allocation_dict:
                del allocation_dict['AC']

            periodic_allocation[period] = allocation_dict

        cy_cp: pd.DataFrame = pd.DataFrame(list(periodic_allocation[end_date].items()), columns=['order_id', 'amount'])
        cy_cp = pd.merge(left=cy_cp, right=dJobs[['order_id', 'customer_code', 'emp_id']], on='order_id', how='left')
        cy_cp_cus: pd.DataFrame = cy_cp.groupby(by='customer_code', as_index=False)['amount'].sum()
        cy_cp_emp: pd.DataFrame = cy_cp.groupby(by='emp_id', as_index=False)['amount'].sum()
        cy_ytd: pd.DataFrame = pd.DataFrame()
        for period in periods:
            month_df: pd.DataFrame = pd.DataFrame(list(periodic_allocation[period].items()),
                                                  columns=['order_id', 'amount'])
            month_df['voucher_date'] = period
            cy_ytd = pd.concat([month_df, cy_ytd])
        cy_ytd = pd.merge(left=cy_ytd, right=dJobs[['order_id', 'customer_code', 'emp_id']], on='order_id',
                          how='left')
        cy_ytd_cus: pd.DataFrame = \
            cy_ytd.groupby(by=[pd.Grouper(key='voucher_date', freq='ME'), 'customer_code'], as_index=False)[
                'amount'].sum()
        cy_ytd_emp: pd.DataFrame = cy_ytd.groupby(by='emp_id', as_index=False)['amount'].sum()
    elif database == 'premium':
        pass
    elif database == 'nbn_logistics':
        pass
    else:
        pass
    return {'periodic_allocation': periodic_allocation, 'cy_cp_cus': cy_cp_cus, 'cy_ytd_cus': cy_ytd_cus,
            'cy_cp_emp': cy_cp_emp, 'cy_ytd_emp': cy_ytd_emp, 'cy_ytd_emp_np': cy_ytd_emp_np,
            'cy_ytd_cus_np': cy_ytd_cus_np, 'cy_cp_cus_np': cy_cp_cus_np, 'cy_cp_emp_np': cy_cp_emp_np,
            'nbnl_profitability': nbnl_profitability}

report = job_profitability(database='elite_security',dEmployee=dEmployee,dExclude=dExclude,dJobs=dJobs,fGL=fGL,fInvoices=fInvoices,fMI=fMI,fOT=fOT,fTimesheet=fTimesheet,
                           end_date=end_date,cogs_map=cogs_ledger_map)

data = report['periodic_allocation'][end_date]
total = sum([v for _,v in data.items()])
print(total)

with open ('periodic_allocation.txt','+a') as file:
    file.write(str(report['periodic_allocation']))

