import pandas as pd
from datetime import datetime

"""
This Programme requires following dataframes
1. fGL
2. fCollection
3. df_LogInv
4. df_dCustomers
"""

START_DATE: datetime = datetime(year=2020, month=11, day=1)
END_DATE: datetime = datetime(year=2024, month=12, day=31)
# VOUCHER_TYPES: list = ['Project Invoice','Contract Invoice','SERVICE INVOICE'] # ESS
VOUCHER_TYPES: list = ['Project Invoice', 'Sales Invoice'] # NBNL

# PATH = r'C:\Masters\Data-ESS.xlsx'
PATH = r'C:\Masters\Data-NBNL.xlsx'


df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
                                    usecols=['voucher_date', 'ledger_code', 'ledger_name', 'transaction_type',
                                             'voucher_number', 'debit', 'forth_level']) # for nbnl
# df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
#                                     usecols=['Voucher Date', 'ledger_code', 'Ledger Name', 'transaction_type',
#                                              'voucher_number', 'Debit Amount', 'Fourth Level Group Name']) # for ess
df_collection: pd.DataFrame = pd.read_excel(io=PATH,
                                            usecols=['ledger_code', 'invoice_number', 'invoice_amount',
                                                     'voucher_number',
                                                     'voucher_date',
                                                     'invoice_date'],
                                            sheet_name='fCollection', date_format={'invoice_date': '%d-%b-%y'},
                                            dtype={
                                                'voucher_number': 'str'},
                                            index_col='invoice_number') # for nbnl

# df_collection: pd.DataFrame = pd.read_excel(io=PATH,
#                                             usecols=['ledger_code', 'Invoice Number', 'invoice_amount',
#                                                      'voucher_number',
#                                                      'voucher_date',
#                                                      'invoice_date'],
#                                             sheet_name='fCollection', date_format={'invoice_date': '%d-%b-%y'},
#                                             dtype={
#                                                 'voucher_number': 'str'},
#                                             index_col='Invoice Number') # for nbnl


df_LogInv :pd.DataFrame = pd.read_excel(io=PATH,usecols=['invoice_date','customer_code'],sheet_name='fLogInv')

# df_dCustomer :pd.DataFrame = pd.read_excel(io=PATH,usecols=['customer_code','Cus_Name','Credit_Days','ledger_code'],sheet_name='dCustomer') # ESS
df_dCustomer :pd.DataFrame = pd.read_excel(io=PATH,usecols=['customer_code','cus_name','credit_days','ledger_code'],sheet_name='dCustomer') # NBNL


def days_taken(row) -> int:
    """For an invoice which has been fully or partially settled calculate how many days it took to settle the invoice in full

        Args:
            row (_type_): a row in df_collection

        Returns:
            int: no of days took to settle an invoice in full
    """
    voucher_dates: list = []
    voucher_amounts: list = []
    voucher_dates_test: list = []
    receipts: list = row['voucher_number'].split(sep=';')
    voucher_amounts = [float(voucher.split(sep='-')[1]) for voucher in receipts]
    if isinstance(row['voucher_date'], datetime):
        voucher_dates.append(row['voucher_date'])
        voucher_dates_test.append(row['voucher_date'] <= END_DATE)
    else:
        date_string :str = row['voucher_date'].split(sep=',')
        voucher_dates = [datetime.strptime(date, '%d-%b-%Y') for date in date_string]
        voucher_dates_test = [date <= END_DATE for date in voucher_dates]

    total_collected :float = sum(amount for amount, valid in zip(voucher_amounts,voucher_dates_test) if valid)
    balance : float = row['invoice_amount'] - total_collected
    if balance == 0:
        return (max(voucher_dates) - row['invoice_date']).days
    else:
        return None


def report(ledger:int, df_collection:pd.DataFrame)->float:
        inv_filt = (df_gl['transaction_type'].isin(VOUCHER_TYPES)) & (df_gl['ledger_code'] == ledger) # for nbnl
        # inv_filt = (df_gl['transaction_type'].isin(VOUCHER_TYPES)) & (df_gl['ledger_code'] == ledger) # for ess
        invoices_list: list = df_gl.loc[inv_filt, 'voucher_number'].unique().tolist()
        # voucher_number, invoices that has not been paid at all
        df_collection: pd.DataFrame = df_collection.loc[df_collection['voucher_number'].notnull()]
        # out of total invoices raised for the whole period, the invoices that were either fully or partially settled. 
        settled_invoices: list = [invoice for invoice in invoices_list if invoice in df_collection.index]
        df_collection = df_collection.loc[
        settled_invoices, ['invoice_amount', 'voucher_number', 'voucher_date', 'invoice_date']]

        if df_collection.empty:
            # for a customer who has not made any payment till date 
            df_collection['days_taken'] = 0
        else:
            df_collection['days_taken'] = df_collection.apply(days_taken, axis=1)
        df_collection = df_collection.loc[df_collection['days_taken'].notnull()]
        df_collection.drop(columns=['voucher_number', 'voucher_date'], inplace=True)
        df_collection.to_csv(f'{ledger}.csv')
        return df_collection['days_taken'].median()

ledgers = [1020201274] # do not enter has str. 
for ledger in ledgers:
    print(report(ledger=ledger,df_collection=df_collection))


def median_days(row)->float:
     ledger_code :int = row['ledger_code'] 
     return report(ledger=ledger_code,df_collection=df_collection)

     
def median_collection_days()->pd.DataFrame:
     start_date :datetime = datetime(year=2024, month=7, day=1)
     end_date :datetime = datetime(year=2024, month=8, day=31)
     inv_filt = (df_LogInv['invoice_date'] >= start_date) & (df_LogInv['invoice_date']<=end_date)
     cust_worked :list = list(set(df_LogInv.loc[inv_filt,'Customer Code'].tolist()))
     report:pd.DataFrame = pd.DataFrame(data={'customer_code':cust_worked})
     report = pd.merge(left=report,right=df_dCustomer,on='customer_code',how='left')
     report['Actual'] = report.apply(median_days,axis=1)
     report = report.fillna(0)
     report.drop(columns=['customer_code','ledger_code'],inplace=True)
     return report


# collection_days = median_collection_days()
# collection_days.to_csv('Collection Days.csv',index=False)