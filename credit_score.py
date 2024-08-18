from datetime import datetime
import pandas as pd
from dateutil.relativedelta import relativedelta
import numpy as np
from od_interest import total_interest


PATH = r'C:\Masters\Data-NBNL.xlsx' # Source of data
# No of months consider for customer to be classified as inactive
OFFSET_MONTHS = relativedelta(months=6)
# Weightage assigned to each parameter
SETTLEMENT_POINTS: int = 30_000  # Weight allocated for time taken to settle invoices in full
AGE_BRACKET_POINTS: int = 20_000  # Weight allocated for each voucher based on their overdue days
FULLY_SETTLED_POINTS: int = AGE_BRACKET_POINTS * 0.25  # Default settlement points for those customers does not have
# receivable balance as on Target date
GP_GENERATED_POINTS: int = 35_000
ESTABLISHED_SINCE_POINTS: int = 5_000  # Weight allocated for the period passed since the incorporation
WORKED_SINCE: int = 10_000
TOTAL: int = SETTLEMENT_POINTS + AGE_BRACKET_POINTS + GP_GENERATED_POINTS + ESTABLISHED_SINCE_POINTS + WORKED_SINCE
OVERDRAFT_INTEREST_PCT :float = 0.08  # Current Overdraft Interest
OVERDRAFT_START_DATE: datetime = datetime(year=2022, month=11, day=13) # Date Overdraft facility started

start_date: datetime = datetime(year=2020, month=11, day=1)
end_date: datetime = datetime(year=2024, month=7, day=31)

df_collection: pd.DataFrame = pd.read_excel(io=PATH,
                                                usecols=['Ledger Code', 'Invoice Number', 'Invoice Amount',
                                                        'Payment Voucher Number',
                                                        'Payment Date',
                                                        'Invoice Date'],
                                                sheet_name='fCollection', date_format={'Invoice Date': '%d-%b-%y'},
                                                dtype={'Payment Voucher Number': 'str'})

df_data: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fData',
                                      usecols=['Voucher_Date', 'Ledger_Code', 'Job_Code', 'Debit', 'Credit'])

df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
                                    usecols=['Voucher Date', 'Amount', 'Ledger Code', 'Ledger Name'])

df_coa: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='dCoAAdler',
                                     usecols=['Ledger_Code', 'Ledger Name', 'Third Level Group Name',
                                              'Second Level Group Name'])

df_customers: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='dCustomers',
                                           usecols=['Customer_Code', 'Cus_Name', 'Ledger_Code', 'Date_Established',
                                                    'Credit_Days'])

df_jobs: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='dJobs', usecols=['Job_Number', 'Customer_Code'])

df_jobs['Job_Code'] = df_jobs['Job_Number'].str.split('-', expand=True)[0].str.strip() # NBNLSIFC240015-Rev1
# |Job_Number|Customer_Code|Job_Code|
df_data['net'] = df_data['Credit'] - df_data['Debit']

df_data = pd.merge(left=df_data[['Voucher_Date', 'Ledger_Code', 'Job_Code', 'net']],
                   right=df_coa[['Ledger_Code', 'Third Level Group Name', 'Ledger Name']], on='Ledger_Code', how='left')
df_data = pd.merge(left=df_data, right=df_jobs[['Job_Code', 'Customer_Code', ]], on='Job_Code', how='left')
df_data = pd.merge(left=df_data, right=df_customers[['Customer_Code', 'Cus_Name']], on='Customer_Code', how='left')


def interest_amount(row)->float:
    """take customer ledger code and return list of jobs for that customer

    Args:
        row (_type_): row in profit report 

    Returns:
        float: list of ledgers for a given customer code
    """
    # a row resembles as follows #|Ledger_Code|Revenue|Profit| and Ledger_Code is the index itself. 
    customer_codes  = df_customers.loc[df_customers['Ledger_Code'] == row.name ,'Customer_Code']
    # as above code generates a series, below will return the corresponding customer_code for a given ledger_code
    customer_code :str = customer_codes.iloc[0]
    # some customer_codes have multiple ledger codes and the duplicates are tagged with "-D". Below will return the customer code as per the system 
    # i.e C00001-D1 and C00001 is a single customer_code comes with different ledger_codes
    customer_code = customer_code.split(sep='-')[0].strip()
    jobs :list = df_jobs.loc[df_jobs['Customer_Code']==customer_code,'Job_Code'].to_list()
    # list of job for a givne customer_code
    inter = total_interest(jobs=jobs)
    return inter


# create a list of dates from start date till end date having first date of each month
monthly_frequencies = pd.date_range(start=start_date, end=end_date, freq='MS')


def profitability_report(df_data:pd.DataFrame,df_gl:pd.DataFrame,df_customers:pd.DataFrame)->pd.DataFrame:

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

        monthly_profit_report = pd.pivot_table(data=df_data_month, index='Cus_Name', values='net', columns='Ledger Name',
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

    profit_report = pd.merge(left=profit_report, right=df_customers[['Cus_Name','Ledger_Code']], how='left', on='Cus_Name')
    profit_report.drop(columns='Cus_Name', inplace=True)
    profit_report = profit_report.groupby('Ledger_Code').sum()
    #|Ledger_Code|Revenue|Profit|
    profit_report['OD_Int'] = profit_report.apply(interest_amount,axis=1)
    profit_report['Profit'] = profit_report['Profit'] - profit_report['OD_Int']
    profit_report.drop(columns='OD_Int',inplace=True)
    return profit_report


filt = ((df_collection['Ledger Code'] < 2000000000) & (~df_collection['Ledger Code'].isin([1020201055])) & (
    df_collection['Invoice Number'].str.contains('NBL/IVL|NBL/PIV|NBL/JV|NBL/CN')) &
        (df_collection['Invoice Date'] <= end_date))

df_collection = df_collection.loc[filt]
# to sort in ascending order, so the last entry in the df will be last transaction of this particular customer
df_collection.sort_values(by=['Ledger Code', 'Invoice Date'], inplace=True)

ledger_code: list = []
first_date: list = []

for i in df_collection['Ledger Code'].unique():
    ledger_code.append(i)
    filt = df_collection['Ledger Code'] == i
    customer_df: pd.DataFrame = df_collection.loc[filt]
    customer_df.reset_index(inplace=True)
    # first date to which a transaction recorded with a particular customer
    first_business_date: datetime = customer_df['Invoice Date'].min()

    for j, _ in customer_df.iterrows():
        # first date of working with a customer is ealiest date of worked with him or if he has a lapse of 6 months
        # between two consecutive transactions dates, then the latest date to which he worked or if the time gap
        # between his last transaction to the end date is more than 6 months, then the end date.
        invoice_date: datetime = customer_df.loc[j, 'Invoice Date']
        last_row: int = len(customer_df) - 1

        if (j != 0) and (j != last_row) and (invoice_date >= (customer_df.loc[j - 1, 'Invoice Date'] + OFFSET_MONTHS)):
            first_business_date = invoice_date
        elif (j != 0) and (j == last_row) and (
                end_date >= (customer_df.loc[last_row, 'Invoice Date'] + OFFSET_MONTHS)) or (j == 0) and (
                end_date >= (customer_df.loc[last_row, 'Invoice Date'] + OFFSET_MONTHS)):
            first_business_date = end_date
        else:
            pass

    first_date.append(first_business_date)
# this will create a datafram which as all the customers in fCollection with their first working date. 
customer_details = pd.DataFrame(data={'Ledger Code': ledger_code, 'First_Date': first_date})
customer_details.set_index(keys='Ledger Code', inplace=True)


def worked_till_brackets(no_of_months: int) -> int:
    """Return points based on the no of months since a customer started working with the company

    Args:
        no_of_months (int): No of months since started working with the company

    Returns:
        int: Points calculated based on no of months since the customer started working 
    """
    if no_of_months <= 12:
        return 1
    if no_of_months <= 24:
        return 2
    if no_of_months <= 36:
        return 3
    if no_of_months <= 48:
        return 4
    return 5


def worked_since_points(customer: int) -> float:
    """Takes Ledger_code of a customer and returns the points based on the no of months a customer have been working 

    Args:
        customer (int): Ledger_code of a customer

    Returns:
        float: Points based on no of months a customer have been working.
    """
    # to get the first date to which the customer has started working with the company
    first_date: datetime = customer_details.loc[customer, 'First_Date']
    period_worked = relativedelta(end_date, first_date)
    # No of months from the date a customer first started working with
    period_worked_months: int = period_worked.months + (period_worked.years * 12) + 1
    return worked_till_brackets(no_of_months=period_worked_months) * WORKED_SINCE / 5


def closing_balance(row) -> float:
    """The function is used to return the remaining balance of each invoice on a given date. Data from fCollection
    sheet is being used. "Payment Voucher Number" has voucher number(s) where each voucher number is separated by
    ";". The amount to which settled by each voucher is shown after "-" [-1] index. "Payment Date" has one or more
    dates which corresponds to the "Payment Voucher Number" and where multiple vouchers numbers exists, individual
    dates are separated by ",".

    Args:
        row (_type_): each row of df_collection dataframe

    Returns:
        float: closing balance of each voucher number on a given date
    """
    if isinstance(row['Payment Voucher Number'],
                  float):  # to handle instances where no receipt/ allocation or credit note issued against an
        # invoice. If no payment made against an invoice,
        # then value of x['Payment Voucher Number'] is float
        return row['Invoice Amount']
    else:
        transactions: list = row['Payment Voucher Number'].split(
            sep=';')  # 'NBL/RV210285-20000.00;NBL/RV210286-3800.00' --> ['NBL/RV210285-20000.00',
        # 'NBL/RV210286-3800.00']
        if ',' in str(row[
                          'Payment Date']):  # to check the existence of multiple payments (then the type will be
            # str) or single payment (then the type will be datetime)
            # and for multiple payments, individual payments are separated by ',' "02-Nov-2020,08-Feb-2021"
            payment_dates: list = [datetime.strptime(date, '%d-%b-%Y') for date in
                                   row['Payment Date'].split(sep=',')]  # 01-Apr-2021,05-Apr-2021 --> [01-Apr-2021,
            # 05-Apr-2021]
            valid_dates: list = [date <= end_date for date in
                                 payment_dates]  # True if the date is less than equal to target date [True,False,True]

            return row['Invoice Amount'] - sum([float(transaction.split(sep='-')[-1]) for transaction in transactions if
                                                valid_dates[transactions.index(transaction)]])  # take the total settled
            # amount of all individual receipts posted against an invoice until a given date and
        # deduct it from the invoice amount which yield the remaining balance of an invoice.  Receipts posted till a
        # target date is True, else False
        else:
            return row['Invoice Amount'] if row['Payment Date'] > end_date else row['Invoice Amount'] - float(
                row['Payment Voucher Number'].split(sep='-')[-1])  # for invoices of which a single receipt posted
            # till today i.e 2/8/2021


def age_points(day_overdue: int) -> int:
    """Based on overdue days for each voucher points will be allocated. Overdue days are after credit period.

    Args:
        day_overdue (int): Overdue days

    Returns:
        int: Points allocated for each voucher based on their overdue no of days. 
    """
    if day_overdue <= 30:
        return 10  # 'Current'
    if day_overdue <= 60:
        return 9  # '0-30'
    if day_overdue <= 90:
        return 8  # '31-60'
    if day_overdue <= 120:
        return 7  # '61-90'
    if day_overdue <= 150:
        return 6  # '91-120'
    if day_overdue <= 180:
        return 5  # '121-150'
    if day_overdue <= 210:
        return 4  # '151-180'
    if day_overdue <= 240:
        return 3  # '181-210'
    if day_overdue <= 270:
        return 2  # '211-240'
    if day_overdue <= 300:
        return 1  # '241-270'
    return 0  # '271+'


def points_for_settlement(days_taken: int) -> int:
    """This function uses days taken to settle the invoice in full, returned by function days_taken_to_settle

    Args:
        output (_type_): No of days taken to settle the invoice after deducting credit period

    Returns:
        _type_: points based on days it took to settle the invoice i.e. Int
    """
    if days_taken <= 30:
        return 10
    if days_taken <= 60:
        return 9
    if days_taken <= 90:
        return 8
    if days_taken <= 120:
        return 7
    if days_taken <= 150:
        return 6
    if days_taken <= 180:
        return 5
    if days_taken <= 210:
        return 4
    if days_taken <= 240:
        return 3
    if days_taken <= 270:
        return 2
    if days_taken <= 300:
        return 1
    return 0


def established_date(customer: int) -> datetime:
    """Takes Ledger_Code assigned to customer and return the established date 

    Args:
        customer (int): Ledger_code of a customer

    Returns:
        datetime: Date established 
    """
    # if customer does not exist in dCustomers, then take the earliest date to which the customer had a transaction
    # in fGL, otherwise take the target date as establishment date
    if pd.isna(cust_master.loc[customer, 'Date_Established']):  # Return True if dCustomer['Date_Established'] is blank
        if customer in df_gl['Ledger Code'].values:  # Check whether target customer exist in df_data
            return df_gl.loc[df_gl['Ledger Code'] == customer, 'Voucher Date'].min()
        else:
            return end_date
    # if customer exist in dCustomers, then the date of establishment is what mentioned under 'Date_Established'
    else:
        return cust_master.loc[customer, 'Date_Established']  # Return date if dCustomer['Date_Established'] has a value


def established_brackets(no_of_months: int) -> int:
    """Takes no of months since the incorporation and return points based on the months

    Args:
        no_of_months (int): no of months passed  since the incorporation

    Returns:
        int: Points based on number of months passed since the incorporation.
    """
    if no_of_months <= (2 * 12):
        return 1
    if no_of_months <= (4 * 12):
        return 2
    if no_of_months <= (6 * 12):
        return 3
    if no_of_months <= (8 * 12):
        return 4
    return 5


def established_points(customer: int) -> float:
    """Total number of points allocated to each customer based on the date of establishment of the company. 

    Args:
        customer (int): Ledger_code of the customer

    Returns:
        float: points allocated for the period since incorporation
    """
    estb_date = established_date(customer=customer)  # Date of establishement for a customer
    years_since = relativedelta(end_date, estb_date)
    # calculate number of months since the incorporation
    period_worked_months: int = years_since.months + (years_since.years * 12) + 1
    # Return points based on the months since the establishment 
    return established_brackets(no_of_months=period_worked_months) * ESTABLISHED_SINCE_POINTS / 5


def age_bracket_points(row) -> float:
    """This returns points to each outstanding invoice based on their outstanding no of days as of target date

    Args:
        row (_type_): Invoice Date as an Input 

    Returns:
        float: Points based on the overdue no of days
    """
    credit_period: int = cust_master.loc[row['Ledger Code'], 'Credit_Days']
    days_lapsed: int = (end_date - row['Invoice Date']).days
    overdue_days: int = days_lapsed - credit_period
    return row['Balance'] * age_points(day_overdue=overdue_days) if row['Balance'] != 0 else 0


def days_taken_to_settle(row) -> float:
    """For those invoices which were settled in full it obtains the latest payment made against an invoice and
    calculate the days it took to settle the invoice in full. Also, it considers the credit period granted for each
    customer. The function uses points_for_settlement to calculate the points based on days it took to settle the
    invoice.

    Args:
        row (_type_): each row of df_collection

    Returns:
        _type_: points calculated based on days it took to settle the invoice i.e. Float
    """
    credit_period: int = cust_master.loc[row['Ledger Code'], 'Credit_Days']  # to obtain the credit period given to
    # each customer
    if row['Balance'] == 0:  # this will consider only those invoices that were fully settled
        if ',' in str(row[
                          'Payment Date']):  # for instances where invoice was settled with multiple payments i.e.
            # '02-Nov-2020,08-Feb-2021'
            total_days: int = (
                    max([datetime.strptime(i, '%d-%b-%Y') for i in row['Payment Date'].split(sep=',')]) - row[
                'Invoice Date']).days  # the difference between the invoice date and the latest date to which the
            # invoice has been settled
            output: int = total_days - credit_period
            # Reduce the credit period of each customer.
        else:
            total_days: int = (row['Payment Date'] - row[
                'Invoice Date']).days  # where invoice has been settled by single transaction
            output: int = total_days - credit_period
        return points_for_settlement(days_taken=output) * row['Invoice Amount']


def thousand_convert(number: int) -> str:
    """take number and convert to 30K format. if the number is 1_000, then return 1K, else number is 1_500 then
    return 1.5K

    Args:
        number (int): number to be converted to 30K format

    Returns:
        str: string with 30K format
    """
    return f"{number // 1_000}K" if number % 1_000 == 0 else f"{number / 1_000:.1f}K"

periodic_profit :pd.DataFrame= profitability_report(df_data=df_data,df_gl=df_gl,df_customers=df_customers)

cust_master: pd.DataFrame = df_customers
cust_master.set_index(keys='Ledger_Code', inplace=True)

df_collection['Balance'] = df_collection.apply(closing_balance, axis=1)
df_collection['Settlement_Points'] = df_collection.apply(days_taken_to_settle, axis=1)
df_collection['Bracket Points'] = df_collection.apply(age_bracket_points, axis=1)

df_collection.drop(columns=['Payment Voucher Number', 'Payment Date'], inplace=True)

# To list of the customers in the chart of accounts
customers: list = df_coa.loc[
    df_coa['Second Level Group Name'].isin(['Due from Related Parties', 'Trade Receivables']), 'Ledger_Code'].to_list()


# empty lists to be filled with calculated points for each customer
settlement_duration: list = []
age_bracket: list = []
gp_generated: list = []
established_since: list = []
worked_since: list = []


for row, customer in enumerate(customers):
    collection_filter = (df_collection['Ledger Code'] == customer) & (df_collection['Balance'] == 0)  # for those
    # customers who do not have outstanding balance on the target date.
    customer_collection_df: pd.DataFrame = df_collection.loc[collection_filter, ['Invoice Amount', 'Settlement_Points']]

    bracket_filter = (df_collection['Ledger Code'] == customer) & (df_collection['Balance'] != 0)  # for those
    # customers who do have a balance on the target date
    bracket_points_df: pd.DataFrame = df_collection.loc[bracket_filter, ['Bracket Points', 'Balance']]

    total_settlement_points: float = customer_collection_df['Settlement_Points'].sum()  # sum of points allocated to
    # each invoice upon full payment
    total_sales: float = customer_collection_df['Invoice Amount'].sum() * 10  # multiplied by 10 as each invoice
    # which has been settled were allocated points from range 0-10

    total_balance: float = bracket_points_df['Balance'].sum() * 10
    bracket_points: float = bracket_points_df['Bracket Points'].sum()  # points allocated to unsettled invoices on the
    # target date based on their overdue days.

    ####Settlement Duration Column Figures Start#####
    if (total_sales == 0) or (np.isnan(total_sales)):
        settlement_points, pct_settlement_points   = (0,0)
    else:
        pct_settlement_points: float = total_settlement_points / total_sales
        settlement_points = pct_settlement_points * SETTLEMENT_POINTS
    settlement_duration.insert(row, settlement_points)
    ####Settlement Duration Column Figures End#####

    ####Age Bracket Column Figures Start#####
    if df_collection.loc[df_collection['Ledger Code'] == customer, 'Balance'].sum() != 0:  # where there is an
        # outstanding balance exist with the customer
        receivable_points: float = bracket_points / total_balance * AGE_BRACKET_POINTS
        age_bracket.insert(row, receivable_points)
    else:
        receivable_points =  pct_settlement_points * FULLY_SETTLED_POINTS  # where the customer does not have
        # outstanding balance on target date
        age_bracket.insert(row, receivable_points)
    ####Age Bracket Column Figures End#####

    ####GP Generated Column Figures Start#####
    if customer in periodic_profit.index:
        revenue: float = periodic_profit.loc[customer, 'Revenue']
        if (revenue == 0) or np.isnan(revenue):
            gp_generated.insert(row, 0)
        else:
            profit_pct: float = periodic_profit.loc[customer, 'Profit'] / revenue
            profit_pct = max(0, min(profit_pct, 1))
            gp_generated.insert(row, profit_pct * GP_GENERATED_POINTS)
    else:
        gp_generated.insert(row, 0)
    ####GP Generated Column Figures end#####

    ####Worked Since Column Figures Start#####
    try:
        worked_since.insert(row, worked_since_points(customer=customer))
    except KeyError:
        worked_since.insert(row, 0)
    ####Worked Since Column Figures end#####

    ####Established Since Column Figures Start#####
    established_since.insert(row, established_points(customer=customer))
    ####Established Since Column Figures End#####

final_report = pd.DataFrame(data={'Customer': customers, 'Settlement Duration': settlement_duration,
                                  'Age Bracket': age_bracket, 'GP Generated': gp_generated,
                                  'Established Since': established_since, 'Worked Since': worked_since})
final_report['Total'] = final_report[
    ['Settlement Duration', 'Age Bracket', 'GP Generated', 'Established Since', 'Worked Since']].sum(axis=1)
final_report.sort_values(by=['Total'], ascending=False, inplace=True)
final_report.rename(columns={'Settlement Duration': f'Settlement Duration({thousand_convert(SETTLEMENT_POINTS)})',
                             'Age Bracket': f'Age Bracket({thousand_convert(AGE_BRACKET_POINTS)})',
                             'GP Generated': f'GP Generated({thousand_convert(GP_GENERATED_POINTS)})',
                             'Established Since': f'Established Since({thousand_convert(ESTABLISHED_SINCE_POINTS)})',
                             'Worked Since': f'Worked Since({thousand_convert(WORKED_SINCE)})',
                             'Total': f'Total({thousand_convert(TOTAL)})',
                             }, inplace=True)

final_report.to_csv(path_or_buf='report.csv', index=False)

# to test NBNLRRI210259 /NBNLAIF240121