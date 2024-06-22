import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Sourse of data
PATH = r'C:\Masters\Data-NBNL.xlsx'

# No of months consider for customer to be classified as inactive
OFFSET_MONTHS = relativedelta(months=6)

# Weightage assigned to each parameters
SETTLEMENT_POINTS = 30_000  # Weight allocated for time taken to settle invoices in full
AGE_BRACKET_POINTS = 20_000  # Weight allocated for each voucher based on their overdue days
FULLY_SETTLED_POINTS = AGE_BRACKET_POINTS * 0.25  # Default settlement points for those customers does not have receivable balance as on Target date
GP_GENERATED = 35_000
ESTABLISHED_SINCE = 5_000
WORKED_SINCE = 10_000

# end_date = input('Please enter closing date yyyy-mm-dd >>')

# end_date = datetime.strptime(end_date, '%Y-%m-%d')
start_date = datetime(year=2020, month=11, day=1)
end_date = datetime(year=2024, month=5, day=31)

df_collection: pd.DataFrame = pd.read_excel(io=PATH,
                                            usecols=['Ledger Code', 'Invoice Number', 'Invoice Amount',
                                                     'Payment Voucher Number',
                                                     'Payment Date',
                                                     'Invoice Date', 'Clear Date'],
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

df_jobs['Job_Code'] = df_jobs['Job_Number'].str.split('-', expand=True)[0].str.strip()
df_data['net'] = df_data['Credit'] - df_data['Debit']

df_data = pd.merge(left=df_data[['Voucher_Date', 'Ledger_Code', 'Job_Code', 'net']],
                   right=df_coa[['Ledger_Code', 'Third Level Group Name', 'Ledger Name']], on='Ledger_Code', how='left')
df_data = pd.merge(left=df_data, right=df_jobs[['Job_Code', 'Customer_Code', ]], on='Job_Code', how='left')
df_data = pd.merge(left=df_data, right=df_customers[['Customer_Code', 'Cus_Name']], on='Customer_Code', how='left')

monthly_frequencies = pd.date_range(start=start_date, end=end_date, freq='MS')

profit_report: pd.DataFrame = pd.DataFrame()

for st_date in monthly_frequencies:
    en_date = st_date + relativedelta(day=31)

    df_data_filt = ((df_data['Voucher_Date'] >= st_date) & (df_data['Voucher_Date'] <= en_date) & (
        df_data['Third Level Group Name'].isin(['Direct Income', 'Cost of Sales'])))

    df_data_month = df_data.loc[df_data_filt]

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
    monthly_profit_report['s_c'] = monthly_profit_report.apply(
        lambda row: row['Logistics Revenue - Clearance'] / rev_clearance * sal_clearance, axis=1)
    monthly_profit_report['s_t'] = monthly_profit_report.apply(
        lambda row: row['Logistics Revenue - Transport'] / rev_transport * sal_transport, axis=1)
    monthly_profit_report['s_f'] = monthly_profit_report.apply(
        lambda row: row['Logistics Revenue - Freight'] / rev_freight * sal_freight, axis=1)
    ###Proportionate allocation of salary cost amoung each transaction###

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

profit_report = pd.merge(left=profit_report, right=df_customers[['Cus_Name', 'Ledger_Code']], how='left', on='Cus_Name')
profit_report.drop(columns='Cus_Name', inplace=True)
profit_report = profit_report.groupby('Ledger_Code').sum()

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
    filt = (df_collection['Ledger Code'] == i)
    customer_df = df_collection.loc[filt]
    # first date to which a transaction recorded with a particular customer
    first_business_date = customer_df['Invoice Date'].min()

    for j, row in customer_df.iterrows():

        invoice_date = row[2]
        # to check whether end_date (i.e report generating date) is 6 months (OFFSET_MONTHS) after the last tranaction done with 
        # current customer under consideration. 
        if end_date >= customer_df.iloc[len(customer_df) - 1, 2] + OFFSET_MONTHS:
            first_business_date = end_date
        elif j == 0 or invoice_date >= (invoice_date - OFFSET_MONTHS):
            continue
        else:
            first_business_date = invoice_date

    first_date.append(first_business_date)
# this will create a datafram which as all the customers in fCollection with their first working date. 
customer_details = pd.DataFrame(data={'Ledger Code': ledger_code, 'First_Date': first_date})
customer_details.set_index(keys='Ledger Code', inplace=True)


def worked_till_brackets(no_of_months):
    if no_of_months <= 12:
        return 1
    elif no_of_months <= 24:
        return 2
    elif no_of_months <= 36:
        return 3
    elif no_of_months <= 48:
        return 4
    else:
        return 5


def worked_since_points(customer):
    first_date = customer_details.loc[customer, 'First_Date']
    period_worked = relativedelta(end_date, first_date)
    period_worked_months = period_worked.months + (period_worked.years * 12) + 1
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
    elif day_overdue <= 60:
        return 9  # '0-30'
    elif day_overdue <= 90:
        return 8  # '31-60'
    elif day_overdue <= 120:
        return 7  # '61-90'
    elif day_overdue <= 150:
        return 6  # '91-120'
    elif day_overdue <= 180:
        return 5  # '121-150'
    elif day_overdue <= 210:
        return 4  # '151-180'
    elif day_overdue <= 240:
        return 3  # '181-210'
    elif day_overdue <= 270:
        return 2  # '211-240'
    elif day_overdue <= 300:
        return 1  # '241-270'
    else:
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
    elif days_taken <= 60:
        return 9
    elif days_taken <= 90:
        return 8
    elif days_taken <= 120:
        return 7
    elif days_taken <= 150:
        return 6
    elif days_taken <= 180:
        return 5
    elif days_taken <= 210:
        return 4
    elif days_taken <= 240:
        return 3
    elif days_taken <= 270:
        return 2
    elif days_taken <= 300:
        return 1
    else:
        return 0


cust_master: pd.DataFrame = df_customers
cust_master.set_index(keys='Ledger_Code', inplace=True)


def established_date(customer):
    if pd.isna(cust_master.loc[customer, 'Date_Established']):  # Return True if dCustomer['Date_Established'] is blank
        if customer in df_gl['Ledger Code'].values:  # Check whether target customer exist in df_data
            return df_gl.loc[df_gl['Ledger Code'] == customer, 'Voucher Date'].min()
        else:
            return end_date
    else:
        return cust_master.loc[customer, 'Date_Established']  # Return date if dCutomer['Date_Established'] has a value


def established_brackets(no_of_months) -> int:
    if no_of_months <= (2 * 12):
        return 1
    elif no_of_months <= (4 * 12):
        return 2
    elif no_of_months <= (6 * 12):
        return 3
    elif no_of_months <= (8 * 12):
        return 4
    else:
        return 5


def established_points(customer):
    estb_date = established_date(customer=customer)
    years_since = relativedelta(end_date, estb_date)
    period_worked_months = years_since.months + (years_since.years * 12) + 1
    return established_brackets(no_of_months=period_worked_months) * ESTABLISHED_SINCE / 5


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


df_collection['Balance'] = df_collection.apply(closing_balance, axis=1)
df_collection['Settlement_Points'] = df_collection.apply(days_taken_to_settle, axis=1)
df_collection['Bracket Points'] = df_collection.apply(age_bracket_points, axis=1)

df_collection.drop(columns=['Payment Voucher Number', 'Payment Date', 'Clear Date'], inplace=True)

customers: list = df_coa.loc[
    df_coa['Second Level Group Name'].isin(['Due from Related Parties', 'Trade Receivables']), 'Ledger_Code'].to_list()
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

    established_since.insert(row, established_points(customer=customer))

    try:
        worked_since.insert(row, worked_since_points(customer=customer))
    except KeyError:
        worked_since.insert(row, 0)

    if customer in profit_report.index:
        try:
            profit_pct: float = profit_report.loc[customer, 'Profit'] / profit_report.loc[customer, 'Revenue']
            profit_pct = max(0, min(profit_pct, 1))
            gp_generated.insert(row, profit_pct * GP_GENERATED)
        except ZeroDivisionError:
            gp_generated.insert(row, 0)
    else:
        gp_generated.insert(row, 0)

    ####Settlement Duration Column Figures Start#####
    try:
        pct_settlement_points: float = total_settlement_points / total_sales
        settlement_points: float = pct_settlement_points * SETTLEMENT_POINTS
        settlement_duration.insert(row, settlement_points)

    except ZeroDivisionError:
        settlement_duration.insert(row, 0)
    ####Settlement Duration Column Figures End#####

    ####Age Bracket Column Figures Start#####
    if df_collection.loc[df_collection['Ledger Code'] == customer, 'Balance'].sum() != 0:  # where there is an
        # outstanding balance exist with the customer
        receivable_points: float = bracket_points / total_balance * AGE_BRACKET_POINTS
        age_bracket.insert(row, receivable_points)
    else:
        receivable_points = pct_settlement_points * FULLY_SETTLED_POINTS  # where the customer does not have
        # outstanding balance on target date
        age_bracket.insert(row, receivable_points)
    ####Age Bracket Column Figures End#####

final_report = pd.DataFrame(data={'Customer': customers, 'Settlement Duration': settlement_duration,
                                  'Age Bracket': age_bracket, 'GP Generated': gp_generated,
                                  'Established Since': established_since, 'Worked Since': worked_since})
final_report['Total'] = final_report[
    ['Settlement Duration', 'Age Bracket', 'GP Generated', 'Established Since', 'Worked Since']].sum(axis=1)
final_report.to_csv(path_or_buf='report.csv', index=False)
