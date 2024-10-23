from datetime import datetime, timedelta
from colorama import Fore, init
import pandas as pd
from dateutil.relativedelta import relativedelta

"""
This Programme requires following dataframes
1. fGL
2. dCustomer
3. fCollection
"""
init()
# PATH = r'C:\Masters\Data-ESS.xlsx'
PATH = r'C:\Masters\Data-NBNL.xlsx'
# Premium Hospitality

START_DATE: datetime = datetime(year=2020, month=11, day=1)
END_DATE: datetime = datetime(year=2024, month=10, day=31)
VOUCHER_TYPES: list = ['Project Invoice', 'Contract Invoice', 'SERVICE INVOICE', 'Sales Invoice']

df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
                                    usecols=['voucher_date', 'ledger_code', 'ledger_name', 'transaction_type',
                                             'voucher_number', 'debit', 'forth_level']) # FOR NBNL
# df_gl: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='fGL',
#                                     usecols=['voucher_date', 'ledger_code', 'ledger_name', 'transaction_type',
#                                              'voucher_number', 'debit', 'forth_level']) # FOR ESS
df_customer: pd.DataFrame = pd.read_excel(io=PATH, sheet_name='dCustomer',
                                          usecols=['ledger_code', 'credit_days'], index_col='ledger_code')
df_collection: pd.DataFrame = pd.read_excel(io=PATH,
                                            usecols=['ledger_code', 'invoice_number', 'invoice_amount',
                                                     'voucher_number',
                                                     'voucher_date',
                                                     'invoice_date'],
                                            sheet_name='fCollection', date_format={'invoice_date': '%d-%b-%y'},
                                            dtype={
                                                'voucher_number': 'str'},
                                            index_col='invoice_number') # FOR NBNL

# df_collection: pd.DataFrame = pd.read_excel(io=PATH,
#                                             usecols=['ledger_code', 'invoice_number', 'Invoice Amount',
#                                                      'voucher_number',
#                                                      'voucher_date',
#                                                      'invoice_date'],
#                                             sheet_name='fCollection', date_format={'invoice_date': '%d-%b-%y'},
#                                             dtype={
#                                                 'voucher_number': 'str'},
#                                             index_col='Invoice Number') # FOR ESS

fLogInv :pd.DataFrame = pd.read_excel(io=PATH,sheet_name='fLogInv',usecols=['invoice_number','emp_id','customer_code'])
# aurang_inv:list = fLogInv.loc[fLogInv['Sales Person Code']=='NBNL0088','Invoice Number'].tolist()
# aurang_inv:list = fLogInv.loc[fLogInv['Customer Code'].isin(['C00174','C00225','CUS0010','CUS0781','CUS0041',
#                                                  'C00137','CUS0012','CUS0013','CUS0610','CUS0613',
#                                                  'C00182','C00210','C00164','C00222','CUS0630',
#                                                  'C00147','C00231','C00136','CUS0787','CUS0792','CUS0794']),'Invoice Number'].tolist()

def receipts_recorded(df_gl: pd.DataFrame, df_collection: pd.DataFrame) -> pd.DataFrame:
    """Takes df_collection and df_gl as arguments. For each voucher that has either fully or partially settled,
    prepares a dataframe which contains voucher number to which the invoice has settled, date to which settled and
    amount to which settled. for those invoices that were settled dataframe consist of. Here voucher number is
    voucher number to which invoice has been settled (i.e. RV/CN/ALN) Voucher Number|voucher_date|credit

    Args:
        df_gl (pd.DataFrame): use to get list of vouchers raised
        df_collection (pd.DataFrame): use to get the history of payments made for a voucher.

    Returns:
        pd.DataFrame: Payment history of each voucher. 
    """
    inv_filt = (df_gl['transaction_type'].isin(VOUCHER_TYPES))
    invoices_list: list = df_gl.loc[inv_filt, 'voucher_number'].unique().tolist()
    # invoices_list: list = [inv for inv in df_gl.loc[inv_filt, 'Voucher Number'].unique().tolist()if inv not in aurang_inv]
    # Payment Voucher Number, invoices that has not been paid at all
    df_collection: pd.DataFrame = df_collection.loc[df_collection['voucher_number'].notnull()]
    # out of total invoices raised for the whole period, the invoices that were either fully or partially settled. 
    settled_invoices: list = [invoice for invoice in invoices_list if invoice in df_collection.index]
    df_collection = df_collection.loc[settled_invoices, ['voucher_number', 'voucher_date']]
    settlement_df_final = pd.DataFrame()
    for invoice in df_collection.index:
        invoice_number: list = []
        voucher_number: list = []
        voucher_date: list = []
        voucher_amount: list = []
        settlement_df = pd.DataFrame()
        # ESS/CN240009-30.08;ESS/RV240204-88403.23 -> ['ESS/CN240009-30.08','ESS/RV240204-88403.23']
        receipts = df_collection.loc[invoice, 'voucher_number'].split(sep=';')
        # ['ESS/CN240009-30.08','ESS/RV240204-88403.23'] -> ['ESS/CN240009','ESS/RV240204']
        voucher_number = [voucher.split(sep='-')[0] for voucher in receipts]
        # ['ESS/CN240009-30.08','ESS/RV240204-88403.23'] -> # [30.08, 88403.23]
        voucher_amount = [float(voucher.split(sep='-')[1]) for voucher in receipts]
        invoice_number = [invoice for _ in range(len(voucher_number))]
        # instance where Payment Date column contains just one date as 5/14/2024
        if isinstance(df_collection.loc[invoice, 'voucher_date'], datetime):
            voucher_date.append(df_collection.loc[invoice, 'voucher_date'])
        else:
            # 31-Mar-2024,05-May-2024 -> [31-Mar-2024,05-May-2024]
            voucher_date = [date for date in df_collection.loc[invoice, 'voucher_date'].split(sep=',')]
        settlement_df: pd.DataFrame = pd.DataFrame(
            data={'voucher_number': voucher_number, 'voucher_date': voucher_date, 'credit': voucher_amount,
                  'invoice_number': invoice_number})
        settlement_df['voucher_date'] = pd.to_datetime(settlement_df['voucher_date'], format='%d-%b-%Y')
        settlement_df_final = pd.concat([settlement_df_final, settlement_df])
        # print(settlement_df_final)
        # print(settlement_df_final.dtypes)
    return settlement_df_final


def closing_date(row) -> datetime:
    """Add credit period (in days) to the voucher date and convert that date to end of the month

    Args:
        row (_type_): a row in dataframe

    Returns:
        datetime: last date of the month to which voucher becomes due
    """
    ledger_code: int = row['ledger_code'] # FOR NBNL
    # ledger_code: int = row['ledger_code'] # FOR ESS
    if ledger_code in df_customer.index:
        credit_days: int = int(df_customer.loc[ledger_code, 'credit_days'])
        due_date = row['voucher_date'] + timedelta(days=credit_days)
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
    due_inv_list: list = list(set(df_already_collected.loc[period_filt, 'voucher_number'].tolist()))
    # due_inv_list: list = [inv for inv in list(set(df_already_collected.loc[period_filt, 'voucher_number'].tolist())) if inv not in aurang_inv]
    collected_filt = (already_collected_receipts['invoice_number'].isin(due_inv_list)) & (
            already_collected_receipts['voucher_date'] < start_date)
    #Uncomment to view already collected amount
    # amount_df  = already_collected_receipts.loc[collected_filt, 'credit']
    # amount_df.to_csv('already_collected.csv')
    amount: float = already_collected_receipts.loc[collected_filt, 'credit'].sum()
    return amount


receipts: pd.DataFrame = receipts_recorded(df_gl=df_gl, df_collection=df_collection)
# receipts.to_csv('collection_report_receipts.csv')
already_collected_receipts: pd.DataFrame = receipts


# filters the collection date based on the selection
filt_collection = (receipts['voucher_date'] >= START_DATE) & (receipts['voucher_date'] <= END_DATE)
receipts = receipts.loc[filt_collection]
# convert collection date to last date of the month, so it can be grouped to know total collected per period.
receipts.loc[:,'voucher_date'] = receipts['voucher_date'].apply(lambda row:row + relativedelta(day=31))
# uncomment below to get the detailed break up of actual collection
# receipts.to_csv('receipts.csv')
receipts = receipts.groupby(by=['voucher_date'], as_index=False)['credit'].sum()
receipts.rename(columns={'voucher_date': 'Due Date', 'credit': 'Actual'}, inplace=True)
# Reasons for Finance / Receipt total for a period not match with 'Actual' in this report
# 1. credit notes are part of 'Actual' in this report
# 2. Receipts other than from customers i.e. Employee Receivable is not part of this report
# 3. Receipts that were not allocated to invoices are not part of this report.
# for 3 above check fCollection/Invoice Number Contains RV/CN and Payment Date ->Blank

filt_net_rev = (df_gl['voucher_date'] >= START_DATE) & (df_gl['voucher_date'] <= END_DATE) & (
    df_gl['transaction_type'].isin(VOUCHER_TYPES)) & (df_gl['forth_level'] == 'Assets')
# filt_net_rev = (df_gl['voucher_Date'] >= START_DATE) & (df_gl['voucher_Date'] <= END_DATE) & (
#     df_gl['transaction_type'].isin(VOUCHER_TYPES)) & (df_gl['forth_level'] == 'Assets') & (~df_gl['voucher_number'].isin(aurang_inv))
df_gl = df_gl.loc[filt_net_rev]
df_gl['Due Date'] = df_gl.apply(closing_date, axis=1)
# Uncommment to see Target amount
# df_gl.to_csv('Target.csv')
df_already_collected: pd.DataFrame = df_gl
df_gl = df_gl.groupby(by=['Due Date'], as_index=False)['debit'].sum()
df_gl['Already_Collected'] = df_gl.apply(already_collected, axis=1)
df_gl['debit'] = df_gl['debit'] - df_gl['Already_Collected']
df_gl.drop(columns=['Already_Collected'], inplace=True)
df_gl = df_gl.loc[(df_gl['Due Date'] >= START_DATE) & (df_gl['Due Date'] <= END_DATE)]
df_gl.rename(columns={'debit': 'Target'}, inplace=True)

combined: pd.DataFrame = pd.concat([receipts.set_index('Due Date'), df_gl.set_index('Due Date')], axis=1,
                                   join='outer').reset_index()
try:
    combined['Performance'] = combined['Actual'] / combined['Target']
except ZeroDivisionError:
    combined['Performance'] = 0

combined.to_csv(path_or_buf='collection_report.csv', index=False)

unallocated: pd.DataFrame = df_collection.loc[(df_collection['invoice_date'] >= datetime(year=2024, month=1, day=1)) & (
    df_collection['voucher_date'].isnull()) & (df_collection.index.str.contains('CN|RV')), ['invoice_date',
                                                                                            'invoice_amount']]
if unallocated.empty:
    print(f'{Fore.GREEN}Good to go.All receipts were allocated{Fore.RESET}')
else:
    print(f'{Fore.RED}Below vouchers requires allocation\n{unallocated.reset_index()}{Fore.RESET}')

