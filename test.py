from dateutil.relativedelta import relativedelta
import pandas as pd
from datetime import datetime,timedelta
import statistics
end_date: datetime = datetime(year=2024, month=7, day=31)

dCustomers_path = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\test_results\dCustomers.csv'
fCollection_path = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\test_results\fCollection.csv'
fInvoices_path = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\test_results\fInvoices.csv'
fGL_path = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\test_results\fGL.csv'

dCustomers:pd.DataFrame = pd.read_csv(dCustomers_path)
fCollection:pd.DataFrame = pd.read_csv(fCollection_path)
fCollection['invoice_date'] = pd.to_datetime(fCollection['invoice_date'])
fCollection['voucher_date'] = pd.to_datetime(fCollection['voucher_date'])
fInvoices:pd.DataFrame = pd.read_csv(fInvoices_path)
fInvoices['Invoice_Date'] = pd.to_datetime(fInvoices['Invoice_Date'])
fGL:pd.DataFrame = pd.read_csv(fGL_path)
fGL['Voucher Date'] = pd.to_datetime(fGL['Voucher Date'])



def settlement_days(invoices: list) -> int:
    col_days: list = []
    invoices = [inv for inv in invoices if not "CN" in inv]
    for invoice in invoices:
        inv_value: float = fCollection.loc[(fCollection['invoice_number'] == invoice), 'invoice_amount'].iloc[0]
        total_collection: float = fCollection.loc[(fCollection['invoice_number'] == invoice) & (
                fCollection['voucher_date'] <= end_date), 'voucher_amount'].sum()
        if (inv_value - total_collection) == 0:
            last_date: datetime = fCollection.loc[(fCollection['invoice_number'] == invoice) & (
                    fCollection['voucher_date'] <= end_date), 'voucher_date'].max()
            inv_date: datetime = fCollection.loc[(fCollection['invoice_number'] == invoice), 'invoice_date'].iloc[0]
            
            col_days.append(last_date - inv_date)
    
    return statistics.median(col_days) if col_days else timedelta(days=0)


def cust_ageing(customer: str) -> pd.DataFrame:
    ledgers: list = dCustomers.loc[(dCustomers['Cus_Name'] == customer), 'Ledger_Code'].tolist()
    credit_days: int = int(dCustomers.loc[dCustomers['Cus_Name'].isin([customer]), 'Credit_Days'].iloc[0])
    invoices: list = list(set(fCollection.loc[fCollection['ledger_code'].isin(ledgers), 'invoice_number'].tolist()))
    cust_soa: pd.DataFrame = fCollection.loc[ (fCollection['invoice_number'].isin(invoices) ), ['invoice_date',
                                                                                                  'invoice_amount',
                                                                                                  'voucher_amount',
                                                                                                  'invoice_number','voucher_date']]
    inv_value_list: list = []
    age_bracket_list: list = []
    ranges = [(0, 'Not Due'), (30, '1-30'), (60, '31-60'),
              (90, '61-90'), (120, '91-120'), (121, '121-150'),
              (151, '151-180'), (181, '181-210'), (211, '211-240'),
              (241, '241-270'), (271, '271-300'), (300, '301-330'),
              (331, '331-360'), (float('inf'), 'More than 361')]
    for invoice in invoices:
        total_collection: float = cust_soa.loc[(cust_soa['invoice_number'] == invoice) & (cust_soa['voucher_date']<=end_date), 'voucher_amount'].sum()
        inv_value: float = cust_soa.loc[(cust_soa['invoice_number'] == invoice), 'invoice_amount'].iloc[0]
        if (inv_value - total_collection) != 0:
            inv_value_list.append(inv_value - total_collection)
            days_passed: int = (end_date - cust_soa.loc[(cust_soa['invoice_number']==invoice),'invoice_date'].iloc[0] - timedelta(days=credit_days)).days
            for threshold, label in ranges:
                if days_passed <= threshold:
                    age_bracket_list.append(label)
                    break
    outstanding_df: pd.DataFrame = pd.DataFrame(
        data={'Inv_Amount': inv_value_list, 'Age Bracket': age_bracket_list}).groupby(by='Age Bracket').sum()
    return outstanding_df


def customer_ratios(customers: list, fInvoices: pd.DataFrame, end_date: datetime, fCollection: pd.DataFrame,
                    dCustomer: pd.DataFrame) -> dict:
    customer_info:dict = {}
    for customer in customers:
        customer_since: datetime = fInvoices.loc[(fInvoices['Cus_Name'] == customer), 'Invoice_Date'].min() if not pd.isna(fInvoices.loc[(fInvoices['Cus_Name'] == customer), 'Invoice_Date'].min()) else "Not Applicable"
        total_revenue: float = fInvoices.loc[
            (fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date), 'Net_Amount'].sum()
        cust_invoices: list = fInvoices.loc[(fInvoices['Cus_Name'] == customer), 'Invoice_Number'].to_list()
        last_receipt_dt: datetime = fCollection.loc[
            fCollection['invoice_number'].isin(cust_invoices), 'voucher_date'].max() if not pd.isna(fCollection.loc[fCollection['invoice_number'].isin(cust_invoices), 'voucher_date'].max()) else "Not Collected"
        print(f'{customer}:{last_receipt_dt}')
        last_receipt_number: str ="Not Collected" if last_receipt_dt =="Not Collected" else fCollection.loc[(fCollection['invoice_number'].isin(cust_invoices)) & (
                fCollection['voucher_date'] == last_receipt_dt), 'voucher_number'].tail(1).iloc[0]
        last_receipt_amt: float ="Not Collected" if last_receipt_dt =="Not Collected" else fCollection.loc[
            (fCollection['voucher_number'] == last_receipt_number), 'voucher_amount'].sum()
        cy_cp_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
                fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=end_date.month,
                                                    day=1)), 'Net_Amount'].sum()
        cy_pp_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
                fInvoices['Invoice_Date'] <= end_date.replace(day=1) - timedelta(days=1)) & (
                                                fInvoices['Invoice_Date'] >= end_date + relativedelta(day=1,
                                                                                                    months=-1)), 'Net_Amount'].sum()
        cy_ytd_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
                fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=1, day=1)), 'Net_Amount'].sum()
        py_ytd_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
                fInvoices['Invoice_Date'] <= datetime(year=end_date.year - 1, month=end_date.month, day=end_date.day)) & (
                                                fInvoices['Invoice_Date'] >= datetime(year=end_date.year - 1, month=1,
                                                                                        day=1)), 'Net_Amount'].sum()
        py_cp_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
                fInvoices['Invoice_Date'] <= datetime(year=end_date.year - 1, month=end_date.month, day=end_date.day)) & (
                                                fInvoices['Invoice_Date'] >= datetime(year=end_date.year - 1,
                                                                                    month=end_date.month,
                                                                                    day=1)), 'Net_Amount'].sum()
        collection_median: float ="Not Collected" if last_receipt_dt =="Not Collected" else settlement_days(invoices=cust_invoices)
        credit_days: int = dCustomer.loc[dCustomers['Cus_Name'].isin([customer]), 'Credit_Days'].iloc[0]
        date_established: datetime = dCustomer.loc[dCustomer['Cus_Name'].isin([customer]), 'Date_Established'].iloc[0]
        outstanding_bal: float = fGL.loc[
            (fGL['Ledger_Code'].isin(dCustomer.loc[dCustomer['Cus_Name'].isin([customer]), 'Ledger_Code'].tolist())) & (
                    fGL['Voucher Date'] <= end_date), 'Amount'].sum()
        cy_cp_rev_contrib_pct: float = fInvoices.loc[
                                        (fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
                                                fInvoices['Invoice_Date'] >= datetime(year=end_date.year,
                                                                                        month=end_date.month,
                                                                                        day=1)), 'Net_Amount'].sum() / \
                                    fInvoices.loc[(fInvoices['Invoice_Date'] <= end_date) & (
                                            fInvoices['Invoice_Date'] >= datetime(year=end_date.year,
                                                                                    month=end_date.month,
                                                                                    day=1)), 'Net_Amount'].sum() * 100
        cy_ytd_rev_contrib_pct: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
                fInvoices['Invoice_Date'] <= end_date) & (fInvoices['Invoice_Date'] >= datetime(year=end_date.year,
                                                                                                month=1,
                                                                                                day=1)), 'Net_Amount'].sum() / \
                                        fInvoices.loc[(fInvoices['Invoice_Date'] <= end_date) & (
                                                fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=1,
                                                                                    day=1)), 'Net_Amount'].sum() * 100
        monthyly_rev: pd.DataFrame = fInvoices.loc[
            (fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
                    fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=1, day=1)), ['Invoice_Date',
                                                                                                'Net_Amount']].groupby(
            by=['Invoice_Date']).sum().rename(columns={'Invoice_Date': 'Month', 'Net_Amount': 'Amount'})
        ageing: pd.DataFrame = cust_ageing(customer=customer)
        last_sales_person:str = fInvoices.loc[(fInvoices['Invoice_Date']<=end_date),'Employee_Code'].tail(1).iloc[0]

        stats: dict = {'customer_since':"Not Applicable" if customer_since =="Not Applicable" else customer_since.strftime('%d-%m-%Y'), 'total_revenue': round(total_revenue), 'credit_score': 0,
                    'last_receipt_amt':"Not Collected" if last_receipt_dt =="Not Collected" else round(last_receipt_amt), 'cy_cp_rev': round(cy_cp_rev), 'cy_pp_rev': round(cy_pp_rev),
                    'last_receipt_dt':"Not Collected" if last_receipt_dt =="Not Collected" else last_receipt_dt.strftime('%d-%m-%Y'),
                    'cy_ytd_rev': round(cy_ytd_rev), 'py_cp_rev': round(py_cp_rev), 'py_ytd_rev': round(py_ytd_rev),
                    'collection_median':"Not Collected" if last_receipt_dt =="Not Collected" else collection_median.days, 'credit_days': credit_days, 'last_sales_person': last_sales_person,
                    'customer_gp': 0, 'outstanding_bal': round(-outstanding_bal), 'ageing': ageing,
                    'date_established': date_established,
                    'cy_cp_rev_contrib_pct': f'{round(cy_cp_rev_contrib_pct,1)}%', 'cy_ytd_rev_contrib_pct': f'{round(cy_ytd_rev_contrib_pct,1)}%',
                    'cy_cp_roi': 0,
                    'cy_ytd_roi': 0, 'monthyly_rev': monthyly_rev, 'remarks': 0}
        customer_info[customer] = stats
    print(customer_info)
    return customer_info


customer_list = dCustomers['Cus_Name'].tolist()

customer_ratios(customers=customer_list,fInvoices=fInvoices,end_date=end_date,fCollection=fCollection,dCustomer=dCustomers)

