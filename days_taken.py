import pandas as pd
from datetime import datetime

"""
This Programme requires following dataframes
1. fGL
2. fCollection
"""

START_DATE: datetime = datetime(year=2020, month=11, day=1)
END_DATE: datetime = datetime(year=2024, month=12, day=31)
VOUCHER_TYPES: list = ['Project Invoice', 'Sales Invoice']
PATH = r'C:\Masters\Data-NBNL.xlsx'

df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
                                    usecols=['Voucher Date', 'Ledger Code', 'Ledger Name', 'Transaction Type',
                                             'Voucher Number', 'Debit Amount', 'Fourth Level Group Name'])
df_collection: pd.DataFrame = pd.read_excel(io=PATH,
                                            usecols=['Ledger Code', 'Invoice Number', 'Invoice Amount',
                                                     'Payment Voucher Number',
                                                     'Payment Date',
                                                     'Invoice Date'],
                                            sheet_name='fCollection', date_format={'Invoice Date': '%d-%b-%y'},
                                            dtype={
                                                'Payment Voucher Number': 'str'},
                                            index_col='Invoice Number')


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


def report(ledger:int,df_collection)->pd.DataFrame:
        inv_filt = (df_gl['Transaction Type'].isin(VOUCHER_TYPES)) & (df_gl['Ledger Code'] == ledger)
        invoices_list: list = df_gl.loc[inv_filt, 'Voucher Number'].unique().tolist()
        # Payment Voucher Number, invoices that has not been paid at all
        df_collection: pd.DataFrame = df_collection.loc[df_collection['Payment Voucher Number'].notnull()]
        # out of total invoices raised for the whole period, the invoices that were either fully or partially settled. 
        settled_invoices: list = [invoice for invoice in invoices_list if invoice in df_collection.index]
        df_collection = df_collection.loc[
        settled_invoices, ['Invoice Amount', 'Payment Voucher Number', 'Payment Date', 'Invoice Date']]
        df_collection['Days_Taken'] = df_collection.apply(days_taken, axis=1)
        df_collection = df_collection.loc[df_collection['Days_Taken'].notnull()]
        df_collection.drop(columns=['Payment Voucher Number', 'Payment Date'], inplace=True)

        df_collection.to_csv(f'{ledger}.csv')

ledgers = [1020201174,1020201325]

for ledger in ledgers:
        report(ledger=ledger,df_collection=df_collection)
