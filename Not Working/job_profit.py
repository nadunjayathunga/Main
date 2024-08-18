from datetime import datetime
import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta

PATH = r'C:\Masters\Data-NBNL.xlsx'
df_data: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fData',
                                      usecols=['Voucher_Date', 'Ledger_Code', 'Job_Code', 'Debit', 'Credit'])
df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
                                    usecols=['Voucher Date', 'Amount', 'Ledger Code', 'Ledger Name'])
df_coa = pd.read_excel(io=PATH, sheet_name='dCoAAdler',
                       usecols=['Ledger_Code', 'Ledger Name', 'Header', 'Fourth Level Group Name',
                                'Third Level Group Name', 'Second Level Group Name'])

df_customers = pd.read_excel(io=PATH, sheet_name='dCustomers', usecols=['Customer_Code', 'Cus_Name','Ledger_Code'])
df_jobs = pd.read_excel(io=PATH, sheet_name='dJobs', usecols=['Job_Number', 'Customer_Code','emp_id'])
df_emp = pd.read_excel(io=PATH, sheet_name='dEmployee', usecols=['Sales_Person_Name', 'emp_id'])

# as jobs numbers may comes in following format NBNLSI244838-Rev2
df_jobs['Job_Code'] = df_jobs['Job_Number'].str.split('-', expand=True)[0].str.strip()
df_data['net'] = df_data['Credit'] - df_data['Debit']

df_data = pd.merge(left=df_data[['Voucher_Date', 'Ledger_Code', 'Job_Code', 'net']],
                   right=df_coa[['Ledger_Code', 'Third Level Group Name', 'Ledger Name']], on='Ledger_Code', how='left')
df_data = pd.merge(left=df_data, right=df_jobs[['Job_Code', 'Customer_Code','emp_id']], on='Job_Code', how='left')
df_data = pd.merge(left=df_data, right=df_customers[['Customer_Code', 'Cus_Name']], on='Customer_Code', how='left')
df_data = pd.merge(left=df_data, right=df_emp, on='emp_id', how='left')

s_date_gl = df_gl['Voucher Date'].min()
s_date_data = df_data['Voucher_Date'].min()
# Maximum of minimum date of GL or Data 
s_date = max(s_date_gl, s_date_data)

e_date_gl = df_gl['Voucher Date'].max()
e_date_data = df_data['Voucher_Date'].max()
# Minimum of maximum of GL or Data
e_date = min(e_date_gl, e_date_data)

report_dict = {'1':'Sales_Person_Name','2':'Cus_Name','3':'Job_Code'}

user_s_date = input(f'Please enter start date yyyy-mm-dd (hint: Min Allowed {s_date.date()}) : ')
user_s_date = datetime.strptime(user_s_date, '%Y-%m-%d')
user_e_date = input(f'Please enter end date yyyy-mm-dd (hint: Max Allowed {e_date.date()}) : ')
user_e_date = datetime.strptime(user_e_date, '%Y-%m-%d')
report_type = input('Please type the report you need\n1.\tEmployee Wise\n2.\tCustomer Wise\n3.\tJob Wise\n>> ')

# list of beginning of the month dates between start date and end date
monthly_frequencies = pd.date_range(start=user_s_date, end=user_e_date, freq='MS')


def profitability_report(df_data:pd.DataFrame,df_gl:pd.DataFrame)->pd.DataFrame:

    # create an empty DataFrame to store Profit made by customers on monthly basis
    profit_report: pd.DataFrame = pd.DataFrame()

    for st_date in monthly_frequencies:
        # To derive the last day of each month
        en_date = st_date + relativedelta(day=31)

        df_data_filt = ((df_data['Voucher_Date'] >= st_date) & (df_data['Voucher_Date'] <= en_date) & (
            df_data['Third Level Group Name'].isin(['Direct Income', 'Cost of Sales'])))

        # Filtered df_data DataFrame contains transactions which have job_id. Since each job_id has a customer,
        # Profitability derived from this DataFrame can be directly apportioned to customers with the exception of direct
        # salary expenes.
        df_data_month: pd.DataFrame = df_data.loc[df_data_filt]

        # pivot the data for a month based on either sales person/ customer name or job code
        monthly_profit_report = pd.pivot_table(data=df_data_month, index=report_dict.get(report_type), values='net', columns='Ledger Name',
                                            aggfunc='sum')
        ###Revenue of each three streams###
        rev_clearance: float = monthly_profit_report['Logistics Revenue - Clearance'].sum()
        rev_transport: float = monthly_profit_report['Logistics Revenue - Transport'].sum()
        rev_freight: float = monthly_profit_report['Logistics Revenue - Freight'].sum()
        ###Revenue of each three streams###

        ###Salary cost of each of three streams###
        sal_clearance: float = df_gl.loc[(df_gl['Voucher Date'] >= st_date) & (df_gl['Voucher Date'] <= en_date) & (
            df_gl['Ledger Name'].isin(
                ['Employee Benefits - Custom Clearance', 'Salaries Expense - Custom Clearance'])), 'Amount'].sum() * -1
        sal_transport: float = df_gl.loc[(df_gl['Voucher Date'] >= st_date) & (df_gl['Voucher Date'] <= en_date) & (
            df_gl['Ledger Name'].isin(
                ['Employee Benefits - Transport', 'Salaries Expense - Transport'])), 'Amount'].sum() * -1
        sal_freight: float = df_gl.loc[(df_gl['Voucher Date'] >= st_date) & (df_gl['Voucher Date'] <= en_date) & (
            df_gl['Ledger Name'].isin(['Employee Benefits - Freight', 'Salaries Expense - Freight'])), 'Amount'].sum() * -1
        ###Salary cost of each of three streams###

        ###Proportionate allocation of salary cost amoung each transaction###
        def calculate_s_c(row:int, rev_clearance:float, sal_clearance:float) -> float:
            """_summary_

            Args:
                row (int): _description_
                rev_clearance (float): _description_
                sal_clearance (float): _description_

            Returns:
                float: _description_
            """
            if rev_clearance == 0:
                return np.nan  # Return NaN if rev_clearance is zero to indicate missing or undefined
            else:
                return row['Logistics Revenue - Clearance'] / rev_clearance * sal_clearance

        def calculate_s_t(row :int, rev_transport :float, sal_transport:float)->float:
            """_summary_

            Args:
                row (int): _description_
                rev_transport (float): _description_
                sal_transport (float): _description_

            Returns:
                float: _description_
            """
            if rev_transport == 0:
                return np.nan  # Return NaN if rev_transport is zero
            else:
                return row['Logistics Revenue - Transport'] / rev_transport * sal_transport

        def calculate_s_f(row:int, rev_freight:float, sal_freight:float) ->float:
            """_summary_

            Args:
                row (int): _description_
                rev_freight (float): _description_
                sal_freight (float): _description_

            Returns:
                float: _description_
            """
            if rev_freight == 0:
                return np.nan  # Return NaN if rev_freight is zero
            else:
                return row['Logistics Revenue - Freight'] / rev_freight * sal_freight
            
        monthly_profit_report['s_c'] = monthly_profit_report.apply(lambda row: calculate_s_c(row, rev_clearance, sal_clearance), axis=1)

        monthly_profit_report['s_t'] = monthly_profit_report.apply(lambda row: calculate_s_t(row, rev_transport, sal_transport), axis=1)

        monthly_profit_report['s_f'] = monthly_profit_report.apply(lambda row: calculate_s_f(row, rev_freight, sal_freight), axis=1)


        # Summing up all revenue ledgers to get the total revenue###
        monthly_profit_report['Revenue'] = monthly_profit_report[
            ['Logistics Revenue - Clearance', 'Logistics Revenue - Freight', 'Logistics Revenue - Transport']].sum(axis=1)
        ###Summing up all cogs ledgers to get the total CoGS###
        monthly_profit_report['Cost'] = monthly_profit_report[
            ['Services Cost - Custom Clearance', 'Services Cost - Freight', 'Services Cost - Transport', 's_c', 's_t',
            's_f']].sum(axis=1)
        monthly_profit_report['Profit'] = monthly_profit_report[['Revenue', 'Cost']].sum(axis=1)

        monthly_profit_report = monthly_profit_report.loc[:, ['Revenue', 'Profit']]
        profit_report = pd.concat([monthly_profit_report, profit_report])
    #|Ledger_Code|Revenue|Profit|
    profit_report = profit_report.groupby(report_dict.get(report_type)).sum()
    profit_report.sort_values(by='Profit', ascending=False, inplace=True)
    return profit_report

report = profitability_report(df_data=df_data,df_gl=df_gl)
report.to_csv(f'job_profit_{user_s_date.date()}-{user_e_date.date()}.csv')

