from datetime import datetime, timedelta

import pandas as pd
from dateutil.relativedelta import relativedelta

PATH = r'C:\Masters\Data-NBNL.xlsx'
# Premium Hospitality

START_DATE: datetime = datetime(year=2020, month=11, day=1)
END_DATE: datetime = datetime(year=2024, month=7, day=1)
VOUCHER_TYPES: list = ['Project Invoice', 'Contract Invoice', 'SERVICE INVOICE', 'Sales Invoice']

df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
                                    usecols=['Voucher Date', 'Ledger Code', 'Ledger Name', 'Transaction Type',
                                             'Voucher Number', 'Debit Amount'])
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
    inv_filt = (df_gl['Transaction Type'].isin(VOUCHER_TYPES))
    invoices_list: list = df_gl.loc[inv_filt, 'Voucher Number'].unique().tolist()
    df_collection: pd.DataFrame = df_collection.loc[df_collection['Payment Voucher Number'].notnull()]
    settled_invoices: list = [invoice for invoice in invoices_list if invoice in df_collection.index]
    df_collection = df_collection.loc[settled_invoices, ['Payment Voucher Number', 'Payment Date']]
    settlement_df_final = pd.DataFrame()
    for invoice in df_collection.index:
        invoice_number :list = []
        voucher_number: list = []
        voucher_date: list = []
        voucher_amount: list = []
        settlement_df = pd.DataFrame()

        receipts = df_collection.loc[invoice, 'Payment Voucher Number'].split(sep=';')
        voucher_number = [voucher.split(sep='-')[0] for voucher in receipts]
        voucher_amount = [float(voucher.split(sep='-')[1]) for voucher in receipts]
        invoice_number.append(invoice)
        if isinstance(df_collection.loc[invoice, 'Payment Date'], datetime):
            voucher_date.append(df_collection.loc[invoice, 'Payment Date'])
        else:
            voucher_date = [date for date in df_collection.loc[invoice, 'Payment Date'].split(sep=',')]
        settlement_df: pd.DataFrame = pd.DataFrame(
            data={'Invoice Number':invoice_number,'Voucher Number': voucher_number, 'Voucher_Date': voucher_date, 'Credit': voucher_amount})
        settlement_df['Voucher_Date'] = pd.to_datetime(settlement_df['Voucher_Date'], format='%d-%b-%Y')
        settlement_df_final = pd.concat([settlement_df_final, settlement_df])
    return settlement_df_final


def closing_date(row) -> datetime:
    ledger_code: int = row['Ledger Code']
    if ledger_code in df_customer.index:
        credit_days: int = int(df_customer.loc[ledger_code, 'Credit_Days'])
        due_date = row['Voucher Date'] + timedelta(days=credit_days)
        return due_date + relativedelta(day=31)
    else:
        pass


def month_end_date(row) -> datetime:
    return row['Voucher_Date'] + relativedelta(day=31)

# def month_collection(row,df_gl:pd.DataFrame)->float:
#     current_month :datetime = row['Due Date']
#     month_invoice_list :list = df_gl.loc[df_gl['Due Date'] == current_month,'Voucher Number'].unique().tolist()


receipts: pd.DataFrame = receipts_recorded(df_gl=df_gl, df_collection=df_collection)
receipts.to_csv('myr.csv')

filt_collection = (receipts['Voucher_Date'] >= START_DATE) & (receipts['Voucher_Date'] <= END_DATE)
receipts = receipts.loc[filt_collection]
receipts['Voucher_Date'] = receipts.apply(month_end_date, axis=1)
receipts = receipts.groupby(by=['Voucher_Date'], as_index=False)['Credit'].sum()
receipts.rename(columns={'Voucher_Date': 'Due Date', 'Credit': 'Actual'}, inplace=True)

filt_net_rev = (df_gl['Voucher Date'] >= START_DATE) & (df_gl['Voucher Date'] <= END_DATE) & (
    df_gl['Transaction Type'].isin(VOUCHER_TYPES))
df_gl = df_gl.loc[filt_net_rev]
df_gl['Due Date'] = df_gl.apply(closing_date, axis=1)
df_gll :pd.DataFrame = df_gl
df_gl.to_csv('myp.csv')
df_gl = df_gl.groupby(by=['Due Date'], as_index=False)['Debit Amount'].sum()
df_gl = df_gl.loc[(df_gl['Due Date'] >= START_DATE) & (df_gl['Due Date'] <= END_DATE)]
df_gl.rename(columns={'Debit Amount': 'Target'}, inplace=True)

combined: pd.DataFrame = pd.concat([receipts.set_index('Due Date'), df_gl.set_index('Due Date')], axis=1,
                                   join='outer').reset_index()
try:
    combined['Performance'] = combined['Actual'] / combined['Target']
except ZeroDivisionError:
    combined['Performance'] = 0

combined.to_csv(path_or_buf='report.csv', index=False)

"""

# from datetime import datetime, timedelta
# import pandas as pd
# from dateutil.relativedelta import relativedelta

# PATH = r'C:\Masters\Data-ESS.xlsx'
# START_DATE = datetime(year=2020, month=11, day=1)
# END_DATE = datetime(year=2024, month=6, day=30)
# VOUCHER_TYPES = ['Project Invoice', 'Contract Invoice', 'SERVICE INVOICE', 'Sales Invoice']

# # Read necessary dataframes
# df_gl = pd.read_excel(io=PATH, sheet_name='fGL',
#                       usecols=['Voucher Date', 'Ledger Code', 'Transaction Type', 'Voucher Number', 'Debit Amount'])

# df_customer = pd.read_excel(io=PATH, sheet_name='dCustomers',
#                             usecols=['Ledger_Code', 'Credit_Days'], index_col='Ledger_Code')

# df_collection = pd.read_excel(io=PATH, sheet_name='fCollection',
#                               usecols=['Invoice Number', 'Payment Voucher Number', 'Payment Date'],
#                               date_format={'Payment Date': '%d-%b-%y'}, dtype={'Payment Voucher Number': 'str'},
#                               index_col='Invoice Number')

# # Filter dataframes based on start and end dates
# df_gl = df_gl[(df_gl['Voucher Date'] >= START_DATE) & (df_gl['Voucher Date'] <= END_DATE)]
# df_collection = df_collection[df_collection['Payment Voucher Number'].notnull()]

# # Function to calculate due dates
# def calculate_due_date(row):
#     ledger_code = row['Ledger Code']
#     if ledger_code in df_customer.index:
#         credit_days = int(df_customer.loc[ledger_code, 'Credit_Days'])
#         due_date = row['Voucher Date'] + timedelta(days=credit_days)
#         return due_date + relativedelta(day=31)
#     else:
#         return None

# # Apply due date calculation to df_gl
# df_gl['Due_Date'] = df_gl.apply(calculate_due_date, axis=1)

# # Aggregate df_gl by due date
# df_gl_aggregated = df_gl.groupby('Due_Date', as_index=False)['Debit Amount'].sum()

# # Function to parse and aggregate receipts
# def aggregate_receipts(df_collection):
#     df_collection['Payment Date'] = pd.to_datetime(df_collection['Payment Date'], format='%d-%b-%Y')
#     receipts = []
#     for _, row in df_collection.iterrows():
#         for voucher in row['Payment Voucher Number'].split(';'):
#             voucher_number, voucher_amount = voucher.split('-')
#             receipts.append({'Voucher Number': voucher_number,
#                              'Voucher Date': row['Payment Date'],
#                              'Credit': float(voucher_amount)})
#     return pd.DataFrame(receipts)

# # Apply receipt aggregation
# receipts_df = aggregate_receipts(df_collection)

# # Filter receipts by date range
# receipts_df = receipts_df[(receipts_df['Voucher Date'] >= START_DATE) & (receipts_df['Voucher Date'] <= END_DATE)]

# # Aggregate receipts by month-end date
# receipts_df['Voucher Date'] = receipts_df['Voucher Date'].apply(lambda x: x + relativedelta(day=31))
# receipts_aggregated = receipts_df.groupby('Voucher Date', as_index=False)['Credit'].sum()

# # Merge aggregated dataframes
# combined_df = pd.merge(receipts_aggregated, df_gl_aggregated, how='outer', left_on='Voucher Date', right_on='Due_Date')

# # Calculate performance
# combined_df['Performance'] = combined_df['Credit'] / combined_df['Debit Amount'].fillna(0)

# # Handle potential division by zero
# combined_df.loc[combined_df['Debit Amount'] == 0, 'Performance'] = 0

# # Rename columns
# combined_df.rename(columns={'Voucher Date': 'Due_Date', 'Credit': 'Actual', 'Debit Amount': 'Target'}, inplace=True)

# # Save to CSV
# combined_df.to_csv('myr.csv', index=False)
    """