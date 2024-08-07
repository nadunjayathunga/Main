from datetime import datetime, timedelta

import pandas as pd
from dateutil.relativedelta import relativedelta

"""
This Programme requires following dataframes
1. fGL
2. dCustomers
3. fCollection
"""

PATH = r'C:\Masters\Data-ESS.xlsx'
# Premium Hospitality

START_DATE: datetime = datetime(year=2020, month=11, day=1)
END_DATE: datetime = datetime(year=2024, month=7, day=31)
VOUCHER_TYPES: list = ['Project Invoice', 'Contract Invoice', 'SERVICE INVOICE', 'Sales Invoice']

df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
                                    usecols=['Voucher Date', 'Ledger Code', 'Ledger Name', 'Transaction Type',
                                             'Voucher Number', 'Debit Amount', 'Fourth Level Group Name'])
df_customer: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='dCustomers',
                                          usecols=['Ledger_Code', 'Credit_Days'], index_col='Ledger_Code')
df_collection: pd.DataFrame = pd.read_excel(io=PATH,
                                            usecols=['Ledger Code', 'Invoice Number', 'Invoice Amount',
                                                     'Payment Voucher Number',
                                                     'Payment Date',
                                                     'Invoice Date'],
                                            sheet_name='fCollection', date_format={'Invoice Date': '%d-%b-%y'},
                                            dtype={
                                                'Payment Voucher Number': 'str'},
                                            index_col='Invoice Number')


def receipts_recorded(df_gl: pd.DataFrame, df_collection: pd.DataFrame) -> pd.DataFrame:
    """Takes df_collection and df_gl as arguments. For each voucher that has either fully or partially settled,
    prepares a dataframe which contains voucher number to which the invoice has settled, date to which settled and
    amount to which settled. for those invoices that were settled dataframe consist of. Here voucher number is
    voucher number to which invoice has been settled (i.e. RV/CN/ALN) Voucher Number|Voucher_Date|Credit

    Args:
        df_gl (pd.DataFrame): use to get list of vouchers raised
        df_collection (pd.DataFrame): use to get the history of payments made for a voucher.

    Returns:
        pd.DataFrame: Payment history of each voucher. 
    """
    inv_filt = (df_gl['Transaction Type'].isin(VOUCHER_TYPES))
    invoices_list: list = df_gl.loc[inv_filt, 'Voucher Number'].unique().tolist()
    # Payment Voucher Number, invoices that has not been paid at all
    df_collection: pd.DataFrame = df_collection.loc[df_collection['Payment Voucher Number'].notnull()]
    # out of total invoices raised for the whole period, the invoices that were either fully or partially settled. 
    settled_invoices: list = [invoice for invoice in invoices_list if invoice in df_collection.index]
    df_collection = df_collection.loc[settled_invoices, ['Payment Voucher Number', 'Payment Date']]
    settlement_df_final = pd.DataFrame()
    for invoice in df_collection.index:
        invoice_number: list = []
        voucher_number: list = []
        voucher_date: list = []
        voucher_amount: list = []
        settlement_df = pd.DataFrame()
        # ESS/CN240009-30.08;ESS/RV240204-88403.23 -> ['ESS/CN240009-30.08','ESS/RV240204-88403.23']
        receipts = df_collection.loc[invoice, 'Payment Voucher Number'].split(sep=';')
        # ['ESS/CN240009-30.08','ESS/RV240204-88403.23'] -> ['ESS/CN240009','ESS/RV240204']
        voucher_number = [voucher.split(sep='-')[0] for voucher in receipts]
        # ['ESS/CN240009-30.08','ESS/RV240204-88403.23'] -> # [30.08, 88403.23]
        voucher_amount = [float(voucher.split(sep='-')[1]) for voucher in receipts]
        invoice_number = [invoice for _ in range(len(voucher_number))]
        # instance where Payment Date column contains just one date as 5/14/2024
        if isinstance(df_collection.loc[invoice, 'Payment Date'], datetime):
            voucher_date.append(df_collection.loc[invoice, 'Payment Date'])
        else:
            # 31-Mar-2024,05-May-2024 -> [31-Mar-2024,05-May-2024]
            voucher_date = [date for date in df_collection.loc[invoice, 'Payment Date'].split(sep=',')]
        settlement_df: pd.DataFrame = pd.DataFrame(
            data={'Voucher Number': voucher_number, 'Voucher_Date': voucher_date, 'Credit': voucher_amount,
                  'Invoice_number': invoice_number})
        settlement_df['Voucher_Date'] = pd.to_datetime(settlement_df['Voucher_Date'], format='%d-%b-%Y')
        settlement_df_final = pd.concat([settlement_df_final, settlement_df])
    return settlement_df_final


def closing_date(row) -> datetime:
    """Add credit period (in days) to the voucher date and convert that date to end of the month

    Args:
        row (_type_): a row in dataframe

    Returns:
        datetime: last date of the month to which voucher becomes due
    """
    ledger_code: int = row['Ledger Code']
    if ledger_code in df_customer.index:
        credit_days: int = int(df_customer.loc[ledger_code, 'Credit_Days'])
        due_date = row['Voucher Date'] + timedelta(days=credit_days)
        return due_date + relativedelta(day=31)
    else:
        pass


def already_collected(row) -> float:
    """Target collection for a given period is calculated by adding the credit period given to each customer.
    Invoices to which Target collection for a given period comprises may contain invoices which has been
    already collected prior they become due or before the beginning of target collection period. i.e. Invoice raised
    in 31/05/2024 which has 60 days credit period will become target collection in the period of 31/07/2024. But if
    such invoice has been collected on 15/06/2024, it should no longer be considered as Target collection for the period
    31/07/2024.

    Args:
        row (_type_): A row in the dataframe

    Returns:
        float: Amount already collected out of target collection
    """
    start_date: datetime = row['Due Date'].replace(day=1)
    period_filt = (df_already_collected['Due Date'] >= start_date) & (
                df_already_collected['Due Date'] <= row['Due Date'])
    due_inv_list: list = list(set(df_already_collected.loc[period_filt, 'Voucher Number'].tolist()))
    collected_filt = (already_collected_receipts['Invoice_number'].isin(due_inv_list)) & (
                already_collected_receipts['Voucher_Date'] < start_date)
    amount: float = already_collected_receipts.loc[collected_filt, 'Credit'].sum()
    return amount


def month_end_date(row) -> datetime:
    """convert voucher date to month end date

    Args:
        row (_type_): a row in a dataframe

    Returns:
        datetime: last date of the month for a given voucher date
    """
    return row['Voucher_Date'] + relativedelta(day=31)


receipts: pd.DataFrame = receipts_recorded(df_gl=df_gl, df_collection=df_collection)
already_collected_receipts: pd.DataFrame = receipts
already_collected_receipts.to_csv('already_collected_receipts.csv')

# filters the collection date based on the selection
filt_collection = (receipts['Voucher_Date'] >= START_DATE) & (receipts['Voucher_Date'] <= END_DATE)
receipts = receipts.loc[filt_collection]
# convert collection date to last date of the month, so it can be grouped to know total collected per period.
receipts['Voucher_Date'] = receipts.apply(month_end_date, axis=1)
receipts = receipts.groupby(by=['Voucher_Date'], as_index=False)['Credit'].sum()
receipts.rename(columns={'Voucher_Date': 'Due Date', 'Credit': 'Actual'}, inplace=True)
# Reasons for Finance / Receipt total for a period not match with 'Actual' in this report
# 1. Credit notes are part of 'Actual' in this report
# 2. Receipts other than from customers i.e. Employee Receivable is not part of this report
# 3. Receipts that were not allocated to invoices are not part of this report.
# for 3 above check fCollection/Invoice Number Contains RV/CN and Payment Date ->Blank

filt_net_rev = (df_gl['Voucher Date'] >= START_DATE) & (df_gl['Voucher Date'] <= END_DATE) & (
    df_gl['Transaction Type'].isin(VOUCHER_TYPES)) & (df_gl['Fourth Level Group Name'] == 'Assets')
df_gl = df_gl.loc[filt_net_rev]
df_gl['Due Date'] = df_gl.apply(closing_date, axis=1)
df_already_collected: pd.DataFrame = df_gl
df_already_collected.to_csv('df_already_collected.csv')
df_gl = df_gl.groupby(by=['Due Date'], as_index=False)['Debit Amount'].sum()
df_gl['Already_Collected'] = df_gl.apply(already_collected, axis=1)
df_gl['Debit Amount'] = df_gl['Debit Amount'] - df_gl['Already_Collected']
df_gl.drop(columns=['Already_Collected'], inplace=True)
df_gl = df_gl.loc[(df_gl['Due Date'] >= START_DATE) & (df_gl['Due Date'] <= END_DATE)]
df_gl.rename(columns={'Debit Amount': 'Target'}, inplace=True)

combined: pd.DataFrame = pd.concat([receipts.set_index('Due Date'), df_gl.set_index('Due Date')], axis=1,
                                   join='outer').reset_index()
try:
    combined['Performance'] = combined['Actual'] / combined['Target']
except ZeroDivisionError:
    combined['Performance'] = 0

combined.to_csv(path_or_buf='collection_report.csv', index=False)
