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
                                    usecols=['Voucher Date', 'Ledger Code', 'Ledger Name', 'Transaction Type',
                                             'Voucher Number', 'Debit Amount', 'Fourth Level Group Name']) # for nbnl
# df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
#                                     usecols=['Voucher Date', 'Ledger_Code', 'Ledger Name', 'Transaction Type',
#                                              'Voucher Number', 'Debit Amount', 'Fourth Level Group Name']) # for ess
df_collection: pd.DataFrame = pd.read_excel(io=PATH,
                                            usecols=['Ledger Code', 'Invoice Number', 'Invoice Amount',
                                                     'Payment Voucher Number',
                                                     'Payment Date',
                                                     'Invoice Date'],
                                            sheet_name='fCollection', date_format={'Invoice Date': '%d-%b-%y'},
                                            dtype={
                                                'Payment Voucher Number': 'str'},
                                            index_col='Invoice Number') # for nbnl

# df_collection: pd.DataFrame = pd.read_excel(io=PATH,
#                                             usecols=['Ledger_Code', 'Invoice Number', 'Invoice Amount',
#                                                      'Payment Voucher Number',
#                                                      'Payment Date',
#                                                      'Invoice Date'],
#                                             sheet_name='fCollection', date_format={'Invoice Date': '%d-%b-%y'},
#                                             dtype={
#                                                 'Payment Voucher Number': 'str'},
#                                             index_col='Invoice Number') # for nbnl


df_LogInv :pd.DataFrame = pd.read_excel(io=PATH,usecols=['Invoice Date','Customer Code'],sheet_name='fLogInv')

# df_dCustomers :pd.DataFrame = pd.read_excel(io=PATH,usecols=['Customer_Code','Cus_Name','Credit_Days','Ledger_Code'],sheet_name='dCustomers') # ESS
df_dCustomers :pd.DataFrame = pd.read_excel(io=PATH,usecols=['Customer_Code','Cus_Name','Credit_Days','Ledger_Code'],sheet_name='dCustomers') # NBNL


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
    receipts: list = row['Payment Voucher Number'].split(sep=';')
    voucher_amounts = [float(voucher.split(sep='-')[1]) for voucher in receipts]
    if isinstance(row['Payment Date'], datetime):
        voucher_dates.append(row['Payment Date'])
        voucher_dates_test.append(row['Payment Date'] <= END_DATE)
    else:
        date_string :str = row['Payment Date'].split(sep=',')
        voucher_dates = [datetime.strptime(date, '%d-%b-%Y') for date in date_string]
        voucher_dates_test = [date <= END_DATE for date in voucher_dates]

    total_collected :float = sum(amount for amount, valid in zip(voucher_amounts,voucher_dates_test) if valid)
    balance : float = row['Invoice Amount'] - total_collected
    if balance == 0:
        return (max(voucher_dates) - row['Invoice Date']).days
    else:
        return None


def report(ledger:int, df_collection:pd.DataFrame)->float:
        inv_filt = (df_gl['Transaction Type'].isin(VOUCHER_TYPES)) & (df_gl['Ledger Code'] == ledger) # for nbnl
        # inv_filt = (df_gl['Transaction Type'].isin(VOUCHER_TYPES)) & (df_gl['Ledger_Code'] == ledger) # for ess
        invoices_list: list = df_gl.loc[inv_filt, 'Voucher Number'].unique().tolist()
        # Payment Voucher Number, invoices that has not been paid at all
        df_collection: pd.DataFrame = df_collection.loc[df_collection['Payment Voucher Number'].notnull()]
        # out of total invoices raised for the whole period, the invoices that were either fully or partially settled. 
        settled_invoices: list = [invoice for invoice in invoices_list if invoice in df_collection.index]
        df_collection = df_collection.loc[
        settled_invoices, ['Invoice Amount', 'Payment Voucher Number', 'Payment Date', 'Invoice Date']]

        if df_collection.empty:
            # for a customer who has not made any payment till date 
            df_collection['Days_Taken'] = 0
        else:
            df_collection['Days_Taken'] = df_collection.apply(days_taken, axis=1)
        df_collection = df_collection.loc[df_collection['Days_Taken'].notnull()]
        df_collection.drop(columns=['Payment Voucher Number', 'Payment Date'], inplace=True)
        df_collection.to_csv(f'{ledger}.csv')
        return df_collection['Days_Taken'].median()

ledgers = [1020201392] # do not enter has str. 
for ledger in ledgers:
    print(report(ledger=ledger,df_collection=df_collection))
     


def median_days(row)->float:
     ledger_code :int = row['Ledger_Code'] 
     return report(ledger=ledger_code,df_collection=df_collection)

     
def median_collection_days()->pd.DataFrame:
     start_date :datetime = datetime(year=2024, month=7, day=1)
     end_date :datetime = datetime(year=2024, month=8, day=31)
     inv_filt = (df_LogInv['Invoice Date'] >= start_date) & (df_LogInv['Invoice Date']<=end_date)
     cust_worked :list = list(set(df_LogInv.loc[inv_filt,'Customer Code'].tolist()))
     report:pd.DataFrame = pd.DataFrame(data={'Customer_Code':cust_worked})
     report = pd.merge(left=report,right=df_dCustomers,on='Customer_Code',how='left')
     report['Actual'] = report.apply(median_days,axis=1)
     report = report.fillna(0)
     report.drop(columns=['Customer_Code','Ledger_Code'],inplace=True)
     return report


# collection_days = median_collection_days()
# collection_days.to_csv('Collection Days.csv',index=False)