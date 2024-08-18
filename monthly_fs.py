import os
from datetime import datetime, timedelta

import pandas as pd
import matplotlib.pyplot as plt
from dateutil.relativedelta import relativedelta
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx2pdf import convert
from io import BytesIO
from data import company_data, company_info
import statistics
import sys

company_id = 0
end_date: datetime = datetime(year=2024, month=7, day=31)
start_date: datetime = datetime(year=end_date.year - 1, month=1, day=1)
VOUCHER_TYPES: list = ['Project Invoice',
                       'Contract Invoice', 'SERVICE INVOICE', 'Sales Invoice']


def data_sources(company_id: int) -> dict:
    path = f'C:\Masters\{company_info[company_id]["data"]["file_name"]}.xlsx'
    fGL: pd.DataFrame = pd.read_excel(io=path, sheet_name='fGL',
                                      usecols=['Bussiness Unit Name', 'Cost Center', 'Voucher Date', 'Credit Amount',
                                               'Debit Amount', 'Narration', 'Ledger Code', 'Voucher Number',
                                               'Transaction Type'])
    dEmployee: pd.DataFrame = pd.read_excel(io=path, sheet_name='dEmployee',
                                            usecols=['Employee_Code',
                                                     'Employee_Name', 'Dept', 'doj', 'nationality', 'Gender',
                                                     'termination_date', 'ba', 'hra', 'tra', 'ma', 'oa', 'pda',
                                                     'travel_cost', 'leave_policy'],
                                            index_col='Employee_Code')
    dCoAAdler: pd.DataFrame = pd.read_excel(io=path, sheet_name='dCoAAdler', index_col='Ledger_Code',
                                            usecols=['Third_Level_Group_Name', 'First_Level_Group_Name', 'Ledger_Code',
                                                     'Ledger_Name', 'Second_Level_Group_Name',
                                                     'Fourth_Level_Group_Name'])
    dCustomers: pd.DataFrame = pd.read_excel(io=path, sheet_name='dCustomers',
                                             usecols=['Customer_Code', 'Ledger_Code', 'Cus_Name', 'Type', 'Credit_Days',
                                                      'Date_Established'])
    fOutSourceInv: pd.DataFrame = pd.read_excel(io=path,
                                                usecols=['Invoice_Number', 'Invoice_Date', 'Customer_Code', 'Job_id',
                                                         'Net_Amount'], sheet_name='fOutSourceInv')
    fAMCInv: pd.DataFrame = pd.read_excel(io=path, sheet_name='fAMCInv',
                                          usecols=['Invoice_Number', 'Invoice_Date', 'Customer_Code', 'Net_Amount','Order_Reference_Number','Sales Engineer Code'])
    fProInv: pd.DataFrame = pd.read_excel(io=path, sheet_name='fProInv',
                                          usecols=['Invoice_Number', 'Invoice_Date', 'Order_ID', 'Customer_Code',
                                                   'Net_Amount','Sales Engineer Code'])
    fCreditNote: pd.DataFrame = pd.read_excel(io=path, sheet_name='fCreditNote',
                                              usecols=['Invoice_Number', 'Invoice_Date', 'Ledger_Code', 'Net_Amount','Job_ID'])

    fBudget: pd.DataFrame = pd.read_excel(io=path,
                                          usecols=['FY', 'L5-Code', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul',
                                                   'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], sheet_name='fBudget')

    fCollection: pd.DataFrame = pd.read_excel(io=path, usecols=['Ledger Code', 'Invoice Number', 'Invoice Amount',
                                                                'Payment Voucher Number', 'Payment Date',
                                                                'Invoice Date'], sheet_name='fCollection',
                                              date_format={
                                                  'Invoice Date': '%d-%b-%y'},
                                              dtype={'Payment Voucher Number': 'str'})
    dCusOrder: pd.DataFrame = pd.read_excel(io=path, usecols=['Order_ID', 'Customer Code', 'Employee_Code'],
                                            sheet_name='dCusOrder')
    dContracts: pd.DataFrame = pd.read_excel(io=path, usecols=['Order_Reference_Number', 'Customer_Code', 'Emp_id'],
                                             sheet_name='dContracts')
    return {'fGL': fGL, 'dEmployee': dEmployee, 'dCoAAdler': dCoAAdler, 'fOutSourceInv': fOutSourceInv,
            'fAMCInv': fAMCInv, 'fProInv': fProInv, 'fCreditNote': fCreditNote, 'dCustomers': dCustomers,
            'fBudget': fBudget, 'fCollection': fCollection, 'dContracts': dContracts, 'dCusOrder': dCusOrder}


def first_page(document, report_date: datetime):
    document.add_picture(
        f'C:\Masters\images\{company_info[company_id]["data"]["abbr"]}-logo.png')
    first = document.add_paragraph()
    first.add_run('\n\n\n')
    first_run = first.add_run(
        company_info[company_id]["data"]["long_name"].upper())
    first_run.bold = True
    first_run.font.size = Pt(24)
    first.alignment = WD_ALIGN_PARAGRAPH.CENTER

    second = document.add_paragraph()
    second_run = second.add_run(
        f'For the period ended {report_date.strftime("%Y-%b-%d")}')
    second_run.font.size = Pt(24)
    second.alignment = WD_ALIGN_PARAGRAPH.CENTER

    third = document.add_paragraph()
    third_run = third.add_run('FINANCIAL STATEMENT ANALYSIS')
    third_run.font.size = Pt(24)
    third.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return document


def business_unit(row, dEmployee: pd.DataFrame) -> str:
    cc: str = row['Cost Center']
    if isinstance(cc, float):
        bu = 'GUARDING-ESS'
    else:
        try:
            if dEmployee.loc[cc, 'Dept'] == 'ELV':
                bu = 'ELV-ESS'
            else:
                bu = 'GUARDING-ESS'
        except:
            bu = 'GUARDING-ESS'
    return bu


def receipts_recorded(data: pd.DataFrame) -> pd.DataFrame:
    data['Payment Date'] = data['Payment Date'].astype('str')
    final_collection_df: pd.DataFrame = pd.DataFrame(columns=[
        'invoice_number', 'ledger_code', 'invoice_date', 'invoice_amount',
        'voucher_number', 'voucher_amount', 'voucher_date'])
    for _, row in data.iterrows():
        pv_number = row['Payment Voucher Number']
        voucher_number = [None] if isinstance(pv_number, float) else [voucher.split(sep='-')[0] for voucher in
                                                                      pv_number.split(sep=';')]
        voucher_amount = [None] if isinstance(pv_number, float) else [float(voucher.split(sep='-')[1]) for voucher in
                                                                      pv_number.split(sep=';')]
        voucher_date = [None] if isinstance(row['Payment Date'], float) else row['Payment Date'].split(sep=',')
        invoice_number = [row['Invoice Number']] if isinstance(pv_number, float) else [row['Invoice Number'] for _ in
                                                                                       range(len(voucher_number))]
        ledger_code = [row['Ledger Code']] if isinstance(pv_number, float) else [row['Ledger Code'] for _ in
                                                                                 range(len(voucher_number))]
        invoice_date = [row['Invoice Date']] if isinstance(pv_number, float) else [row['Invoice Date'] for _ in
                                                                                   range(len(voucher_number))]
        invoice_amount = [row['Invoice Amount']] if isinstance(pv_number, float) else [row['Invoice Amount'] for _ in
                                                                                       range(len(voucher_number))]
        collection_df: pd.DataFrame = pd.DataFrame(
            data={'invoice_number': invoice_number, 'ledger_code': ledger_code, 'invoice_date': invoice_date,
                  'invoice_amount': invoice_amount,
                  'voucher_number': voucher_number, 'voucher_amount': voucher_amount, 'voucher_date': voucher_date})
        final_collection_df = pd.concat([final_collection_df, collection_df])
    return final_collection_df


def preprocessing(data: dict) -> dict:
    fGL: pd.DataFrame = data['fGL']
    dEmployee: pd.DataFrame = data['dEmployee']
    dCoAAdler: pd.DataFrame = data['dCoAAdler']
    fOutSourceInv: pd.DataFrame = data['fOutSourceInv']
    fAMCInv: pd.DataFrame = data['fAMCInv']
    fProInv: pd.DataFrame = data['fProInv']
    fCreditNote: pd.DataFrame = data['fCreditNote']
    dCustomers: pd.DataFrame = data['dCustomers']
    fBudget: pd.DataFrame = data['fBudget']
    fCollection: pd.DataFrame = data['fCollection']
    dContracts: pd.DataFrame = data['dContracts']
    dCusOrder: pd.DataFrame = data['dCusOrder']

    fGL['Cost Center'] = fGL['Cost Center'].str.split(
        '|', expand=True)[0].str.strip()  # ESS0012 | GAURAV VASHISTH
    fGL['Bussiness Unit Name'] = fGL.apply(
        business_unit, axis=1, args=[dEmployee])
    fGL.replace(
        to_replace={'Elite Security Services': 'GUARDING-ESS'}, inplace=True)
    fGL['Amount'] = fGL['Credit Amount'] - fGL['Debit Amount']
    fGL.drop(columns=['Credit Amount', 'Debit Amount'], inplace=True)
    fGL['Voucher Date'] = fGL.apply(
        lambda row: row['Voucher Date'] + relativedelta(day=31), axis=1)
    fGL.rename(columns={'Ledger Code': 'Ledger_Code'}, inplace=True)

    fOutSourceInv = pd.merge(
        left=fOutSourceInv, right=dCustomers, on='Customer_Code', how='left')
    fAMCInv = pd.merge(left=fAMCInv, right=dCustomers,
                       on='Customer_Code', how='left')
    fProInv = pd.merge(left=fProInv, right=dCustomers,
                       on='Customer_Code', how='left')
    fCreditNote['Net_Amount'] = fCreditNote['Net_Amount'] * -1
    fCreditNote = pd.merge(
        left=fCreditNote, right=dCustomers, on='Ledger_Code', how='left')

    fInvoices: pd.DataFrame = pd.concat(
        [fOutSourceInv, fAMCInv, fProInv, fCreditNote])
    fInvoices.drop(columns=['Order_ID'], inplace=True)

    fBudget = pd.melt(fBudget, id_vars=[
        'FY', 'L5-Code'], var_name='Month', value_name='Amount')
    fBudget.rename(columns={'L5-Code': 'Ledger_Code'}, inplace=True)
    fBudget['Voucher Date'] = fBudget.apply(
        lambda x: pd.to_datetime(f'{x["FY"]}-{x["Month"]}-01') + relativedelta(day=31), axis=1)
    fBudget.drop(columns=['FY', 'Month'], inplace=True)
    fBudget = fBudget.loc[fBudget['Amount'] != 0]
    fBudget = pd.merge(left=fBudget, right=dCoAAdler,
                       on='Ledger_Code', how='left')
    fBudget['Bussiness Unit Name'] = 'GUARDING-ESS'
    fCollection = receipts_recorded(data=fCollection)

    dContracts.rename(columns={'Order_Reference_Number': 'Order_ID', 'Emp_id': 'Employee_Code'}, inplace=True)
    dContracts['Order_ID'] = dContracts['Order_ID'].str.split('-', expand=True)[0]
    dJobs: pd.DataFrame = pd.concat([dContracts, dCusOrder], ignore_index=True)

    return {'fGL': fGL, 'dEmployee': dEmployee, 'dCoAAdler': dCoAAdler, 'fInvoices': fInvoices, 'fBudget': fBudget,
            'dCustomers': dCustomers, 'fCollection': fCollection, 'dJobs': dJobs}


def empctc(row, dEmployee: pd.DataFrame) -> float:
    policy: str = dEmployee.loc[(dEmployee['Employee_Code'] == row['Employee_Code']), 'leave_policy']
    basic: float = dEmployee.loc[(dEmployee['Employee_Code'] == row['Employee_Code']), 'ba']
    gross: float = \
        dEmployee.loc[
            (dEmployee['Employee_Code'] == row['Employee_Code']), ['ba', 'hra', 'tra', 'ma', 'oa', 'pda']].sum(
            axis=1).values[0]
    ticket_amt: float = dEmployee.loc[(dEmployee['Employee_Code'] == row['Employee_Code']), 'travel_cost']
    try:
        leave_policy, ticket_policy = int(policy.split(sep='-')[1].strip().split(sep=' ')[0]), \
            policy.split(sep='-')[-1].strip().split(sep=' ')[0]
    except IndexError:
        leave_policy, ticket_policy = 1, 1
    eos: float = basic * 12 / 365 * 30 / 12
    leave: float = gross * 12 / 365 * leave_policy / 12
    ticket: float = ticket_amt / (12 if ticket_policy == 'Yearly' else 24)
    ctc: float = gross + eos + leave + ticket
    return ctc


cleaned_data: dict = preprocessing(data=data_sources(company_id=0))
fGL: pd.DataFrame = cleaned_data['fGL']
dEmployee: pd.DataFrame = cleaned_data['dEmployee']
dCoAAdler: pd.DataFrame = cleaned_data['dCoAAdler']
merged: pd.DataFrame = pd.merge(
    left=fGL, right=dCoAAdler, on='Ledger_Code', how='left')
fInvoices: pd.DataFrame = cleaned_data['fInvoices']
fBudget: pd.DataFrame = cleaned_data['fBudget']
fCollection: pd.DataFrame = cleaned_data['fCollection']
dCustomers: pd.DataFrame = cleaned_data['dCustomers']
dJobs: pd.DataFrame = cleaned_data['dJobs']


def profitandlossheads(data: pd.DataFrame, start_date: datetime, end_date: datetime, bu: list) -> pd.DataFrame:
    gp_filt = (data['Third_Level_Group_Name'].isin(['Cost of Sales', 'Direct Income'])) & (
            data['Voucher Date'] >= start_date) & (data['Voucher Date'] <= end_date) & (
                  data['Bussiness Unit Name'].isin(bu))
    gp: float = data.loc[gp_filt, 'Amount'].sum()
    oh_filt = (data['Third_Level_Group_Name'].isin(['Overhead', 'Finance Cost'])) & (
            data['Voucher Date'] >= start_date) & (data['Voucher Date'] <= end_date) & (
                  data['Bussiness Unit Name'].isin(bu))
    overhead: float = data.loc[oh_filt, 'Amount'].sum()
    np_filt = (data['Fourth_Level_Group_Name'].isin(['Expenses', 'Income'])) & (data['Voucher Date'] >= start_date) & (
            data['Voucher Date'] <= end_date) & (data['Bussiness Unit Name'].isin(bu))
    np: float = data.loc[np_filt, 'Amount'].sum()
    rev_filt = (data['Third_Level_Group_Name'].isin(['Direct Income'])) & (
            data['Voucher Date'] >= start_date) & (data['Voucher Date'] <= end_date) & (
                   data['Bussiness Unit Name'].isin(bu))
    rev: float = data.loc[rev_filt, 'Amount'].sum()
    gp_row: pd.DataFrame = pd.DataFrame(data={'Amount': [gp], 'Voucher Date': [end_date]}, index=['Gross Profit'])
    oh_row: pd.DataFrame = pd.DataFrame(data={'Amount': [overhead], 'Voucher Date': [end_date]},
                                        index=['Total Overhead'])
    np_row: pd.DataFrame = pd.DataFrame(data={'Amount': [np], 'Voucher Date': [end_date]}, index=['Net Profit'])
    rev_row: pd.DataFrame = pd.DataFrame(data={'Amount': [rev], 'Voucher Date': [end_date]}, index=['Total Revenue'])
    pl_summary: pd.DataFrame = pd.concat([gp_row, oh_row, np_row, rev_row], ignore_index=False)
    return pl_summary


def profitandloss(data: pd.DataFrame, start_date: datetime, end_date: datetime, basic_pl: bool = False,
                  mid_pl: bool = False, full_pl: bool = False,
                  bu: list = list(set(fGL['Bussiness Unit Name'].tolist()))) -> dict:
    df_basic: pd.DataFrame = pd.DataFrame(data={'Voucher Date': [], 'Amount': []})
    df_basic_bud: pd.DataFrame = pd.DataFrame(data={'Voucher Date': [], 'Amount': []})
    df_mid: pd.DataFrame = pd.DataFrame(data={'Voucher Date': [], 'Amount': []})
    df_mid_bud: pd.DataFrame = pd.DataFrame(data={'Voucher Date': [], 'Amount': []})
    df_full: pd.DataFrame = pd.DataFrame(data={'Voucher Date': [], 'Amount': []})
    df_full_bud: pd.DataFrame = pd.DataFrame(data={'Voucher Date': [], 'Amount': []})
    basic: pd.DataFrame = pd.DataFrame()
    basic_bud: pd.DataFrame = pd.DataFrame()
    mid: pd.DataFrame = pd.DataFrame()
    mid_bud: pd.DataFrame = pd.DataFrame()
    full: pd.DataFrame = pd.DataFrame()
    full_bud: pd.DataFrame = pd.DataFrame()
    month_end_dates = pd.date_range(start=start_date, end=end_date, freq='M')
    for end in month_end_dates:
        start: datetime = end + relativedelta(day=1)
        indirect_inc_filt = data['Third_Level_Group_Name'].isin(['Indirect Income']) & (
                data['Voucher Date'] >= start) & (data['Voucher Date'] <= end) & (
                                data['Bussiness Unit Name'].isin(bu))
        indirect_inc_brief: pd.DataFrame = data.loc[
            indirect_inc_filt, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
            by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
            columns={'First_Level_Group_Name': 'Description'})
        indirect_inc_brief = indirect_inc_brief.loc[indirect_inc_brief['Amount'] != 0]

        indirect_inc_filt_bud = fBudget['Third_Level_Group_Name'].isin(['Indirect Income']) & (
                fBudget['Voucher Date'] >= start) & (fBudget['Voucher Date'] <= end) & (
                                    fBudget['Bussiness Unit Name'].isin(bu))
        indirect_inc_brief_bud: pd.DataFrame = fBudget.loc[
            indirect_inc_filt_bud, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
            by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
            columns={'First_Level_Group_Name': 'Description'})
        indirect_inc_brief_bud = indirect_inc_brief_bud.loc[indirect_inc_brief_bud['Amount'] != 0]

        overhead_brief_filt = data['Third_Level_Group_Name'].isin(['Overhead', 'Finance Cost']) & (
                data['Voucher Date'] >= start) & (data['Voucher Date'] <= end) & (
                                  data['Bussiness Unit Name'].isin(bu))
        overhead_brief_filt_bud = fBudget['Third_Level_Group_Name'].isin(['Overhead', 'Finance Cost']) & (
                fBudget['Voucher Date'] >= start) & (data['Voucher Date'] <= end) & (
                                      fBudget['Bussiness Unit Name'].isin(bu))
        summary_actual: pd.DataFrame = profitandlossheads(data=data, start_date=start, end_date=end, bu=bu)
        summary_budget: pd.DataFrame = profitandlossheads(data=fBudget, start_date=start, end_date=end, bu=bu)
        # basic version
        if basic_pl:
            trade_account_filt = data['Third_Level_Group_Name'].isin(['Cost of Sales', 'Direct Income']) & (
                    data['Voucher Date'] >= start) & (data['Voucher Date'] <= end) & (
                                     data['Bussiness Unit Name'].isin(bu))
            trade_account_filt_bud = fBudget['Third_Level_Group_Name'].isin(['Cost of Sales', 'Direct Income']) & (
                    fBudget['Voucher Date'] >= start) & (fBudget['Voucher Date'] <= end) & (
                                         fBudget['Bussiness Unit Name'].isin(bu))
            overhead_brief_basic: pd.DataFrame = data.loc[
                overhead_brief_filt, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
                columns={'First_Level_Group_Name': 'Description'})
            overhead_brief_basic_bud: pd.DataFrame = fBudget.loc[
                overhead_brief_filt_bud, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
                columns={'First_Level_Group_Name': 'Description'}, inplace=True)
            trade_account_brief: pd.DataFrame = data.loc[
                trade_account_filt, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
                columns={'First_Level_Group_Name': 'Description'})
            trade_account_brief_bud: pd.DataFrame = fBudget.loc[
                trade_account_filt_bud, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
                columns={'First_Level_Group_Name': 'Description'})
            basic: pd.DataFrame = pd.concat(
                [trade_account_brief, indirect_inc_brief, overhead_brief_basic]).rename(
                columns={'First_Level_Group_Name': 'Description'})
            basic_bud: pd.DataFrame = pd.concat(
                [trade_account_brief_bud, indirect_inc_brief_bud, overhead_brief_basic_bud]).rename(
                columns={'First_Level_Group_Name': 'Description'})
            basic = basic.loc[basic['Amount'] != 0].set_index(keys='Description')
            basic_bud = basic_bud.loc[basic_bud['Amount'] != 0].set_index(keys='Description')
            # if not [df for df in [basic, summary_actual, df_basic] if df.empty]:
            #     print([df for df in [basic, summary_actual, df_basic] if df.empty])
            #     print(f'start:{start},end:{end}')
            df_basic = pd.concat([basic, summary_actual, df_basic])
            df_basic_bud = pd.concat([basic_bud, summary_budget, df_basic_bud])

        # mid version
        if mid_pl:
            trade_account_mid: pd.DataFrame = data.loc[
                trade_account_filt, ['Second_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'Second_Level_Group_Name'], as_index=False).sum().rename(
                columns={'Second_Level_Group_Name': 'Description'})
            trade_account_mid_bud: pd.DataFrame = fBudget.loc[
                trade_account_filt_bud, ['Second_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'Second_Level_Group_Name'], as_index=False).sum().rename(
                columns={'Second_Level_Group_Name': 'Description'})
            overhead_brief_mid: pd.DataFrame = data.loc[
                overhead_brief_filt, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
                columns={'First_Level_Group_Name': 'Description'})
            overhead_brief_mid_bud: pd.DataFrame = fBudget.loc[
                overhead_brief_filt_bud, ['First_Level_Group_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'First_Level_Group_Name'], as_index=False).sum().rename(
                columns={'First_Level_Group_Name': 'Description'})
            mid = pd.concat(
                [trade_account_mid, indirect_inc_brief, overhead_brief_mid])
            mid_bud = pd.concat(
                [trade_account_mid_bud, indirect_inc_brief_bud, overhead_brief_mid_bud])
            mid = mid.loc[mid['Amount'] != 0].set_index(keys='Description')
            mid_bud = mid_bud.loc[mid_bud['Amount'] != 0].set_index(keys='Description')
            df_mid = pd.concat([mid, summary_actual, df_mid])
            df_mid_bud = pd.concat([mid_bud, summary_budget, df_mid_bud])

        # full version
        if full_pl:
            detailed_filt = data['Third_Level_Group_Name'].isin(
                ['Indirect Income', 'Overhead', 'Finance Cost', 'Direct Income', 'Cost of Sales']) & (
                                    data['Voucher Date'] >= start) & (data['Voucher Date'] <= end) & (
                                data['Bussiness Unit Name'].isin(bu))
            detailed_filt_bud = fBudget['Third_Level_Group_Name'].isin(
                ['Indirect Income', 'Overhead', 'Finance Cost', 'Direct Income', 'Cost of Sales']) & (
                                        fBudget['Voucher Date'] >= start) & (fBudget['Voucher Date'] <= end) & (
                                    fBudget['Bussiness Unit Name'].isin(bu))
            full = data.loc[detailed_filt, ['Ledger_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'Ledger_Name'], as_index=False).sum().rename(columns={'Ledger_Name': 'Description'})
            full_bud = fBudget.loc[detailed_filt_bud, ['Ledger_Name', 'Voucher Date', 'Amount']].groupby(
                by=['Voucher Date', 'Ledger_Name'], as_index=False).sum().rename(columns={'Ledger_Name': 'Description'})
            full = full.loc[full['Amount'] != 0].set_index(keys='Description')
            full_bud = full_bud.loc[full_bud['Amount'] != 0].set_index(keys='Description')
            df_full = pd.concat([df_full, summary_actual, full])
            df_full_bud = pd.concat([df_full_bud, summary_budget, full_bud])

    cy_cp_basic: pd.DataFrame = df_basic.loc[df_basic['Voucher Date'] == end_date].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})
    cy_cp_basic_bud: pd.DataFrame = df_basic_bud.loc[df_basic_bud['Voucher Date'] == end_date].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    cy_pp_basic: pd.DataFrame = df_basic.loc[
        df_basic['Voucher Date'] == end_date - relativedelta(months=1) + relativedelta(day=31)].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    cy_ytd_basic: pd.DataFrame = df_basic.loc[(df_basic['Voucher Date'] <= end_date) & (
            df_basic['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()
    cy_ytd_basic_bud: pd.DataFrame = df_basic_bud.loc[(df_basic_bud['Voucher Date'] <= end_date) & (
            df_basic_bud['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()

    py_cp_basic: pd.DataFrame = df_basic.loc[
        df_basic['Voucher Date'] == datetime(year=end_date.year - 1, month=end_date.month, day=end_date.day)].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    py_ytd_basic: pd.DataFrame = df_basic.loc[
        (df_basic['Voucher Date'] <= end_date - relativedelta(years=1) + relativedelta(day=31)) & (
                df_basic['Voucher Date'] >= datetime(year=end_date.year - 1, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()

    cy_ytd_basic_monthwise: pd.DataFrame = df_basic.loc[(df_basic['Voucher Date'] <= end_date) & (
            df_basic['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).pivot_table(index='Description', columns='Voucher Date', values='Amount',
                                                      aggfunc='sum', fill_value=0).reset_index()

    cy_cp_mid: pd.DataFrame = df_mid.loc[df_mid['Voucher Date'] == end_date].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})
    cy_cp_mid_bud: pd.DataFrame = df_mid_bud.loc[df_mid['Voucher Date'] == end_date].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    cy_pp_mid: pd.DataFrame = df_mid.loc[
        df_mid['Voucher Date'] == end_date - relativedelta(months=1) + relativedelta(day=31)].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    cy_ytd_mid: pd.DataFrame = df_mid.loc[
        (df_mid['Voucher Date'] <= end_date) & (
                df_mid['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()
    cy_ytd_mid_bud: pd.DataFrame = df_mid_bud.loc[
        (df_mid_bud['Voucher Date'] <= end_date) & (
                df_mid_bud['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()

    py_cp_mid: pd.DataFrame = df_mid.loc[
        df_mid['Voucher Date'] == datetime(year=end_date.year - 1, month=end_date.month, day=end_date.day)].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    py_ytd_mid: pd.DataFrame = df_mid.loc[
        (df_mid['Voucher Date'] <= end_date - relativedelta(years=1) + relativedelta(day=31)) & (
                df_mid['Voucher Date'] >= datetime(year=end_date.year - 1, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()

    cy_cp_full: pd.DataFrame = df_full.loc[df_full['Voucher Date'] == end_date].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})
    cy_cp_full_bud: pd.DataFrame = df_full_bud.loc[df_full_bud['Voucher Date'] == end_date].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    cy_pp_full: pd.DataFrame = df_full.loc[
        df_full['Voucher Date'] == end_date - relativedelta(months=1) + relativedelta(day=31)].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    cy_ytd_full: pd.DataFrame = df_full.loc[(df_full['Voucher Date'] <= end_date) & (
            df_full['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()
    cy_ytd_full_bud: pd.DataFrame = df_full_bud.loc[(df_full_bud['Voucher Date'] <= end_date) & (
            df_full_bud['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()

    py_cp_full: pd.DataFrame = df_full.loc[
        df_full['Voucher Date'] == datetime(year=end_date.year - 1, month=end_date.month, day=end_date.day)].drop(
        columns=['Voucher Date']).reset_index().rename(columns={'index': 'Description'})

    py_ytd_full: pd.DataFrame = df_full.loc[
        (df_full['Voucher Date'] <= end_date - relativedelta(years=1) + relativedelta(day=31)) & (
                df_full['Voucher Date'] >= datetime(year=end_date.year - 1, month=1, day=1))].reset_index().rename(
        columns={'index': 'Description'}).drop(columns=['Voucher Date']).groupby(by=['Description'],
                                                                                 as_index=False).sum()

    return {'df_basic': {'cy_cp_basic': cy_cp_basic, 'cy_pp_basic': cy_pp_basic, 'cy_ytd_basic': cy_ytd_basic,
                         'py_cp_basic': py_cp_basic, 'py_ytd_basic': py_ytd_basic, 'cy_cp_basic_bud': cy_cp_basic_bud,
                         'cy_ytd_basic_bud': cy_ytd_basic_bud, 'cy_ytd_basic_monthwise': cy_ytd_basic_monthwise},

            'df_mid': {'cy_cp_mid': cy_cp_mid, 'cy_pp_mid': cy_pp_mid, 'cy_ytd_mid': cy_ytd_mid, 'py_cp_mid': py_cp_mid,
                       'py_ytd_mid': py_ytd_mid, 'cy_cp_mid_bud': cy_cp_mid_bud, 'cy_ytd_mid_bud': cy_ytd_mid_bud},

            'df_full': {'cy_cp_full': cy_cp_full, 'cy_pp_full': cy_pp_full, 'cy_ytd_full': cy_ytd_full,
                        'py_cp_full': py_cp_full, 'py_ytd_full': py_ytd_full, 'cy_cp_full_bud': cy_cp_full_bud,
                        'cy_ytd_full_bud': cy_ytd_full_bud}}


def offset_bal(end_date: datetime, data: pd.DataFrame) -> float:
    offset_accounts: list = company_info[0]['data']['offset_accounts']
    offset_filt = (data['Voucher Date'] <= end_date) & (
        data['Ledger_Code'].isin(offset_accounts))
    offset_df: float = data.loc[offset_filt, 'Amount'].sum()
    return offset_df


def interco_bal(data: pd.DataFrame, end_date: datetime) -> dict:
    interco_final: pd.DataFrame = pd.DataFrame()
    for entity in company_data:
        interco_ids: list = dCoAAdler.loc[dCoAAdler['Ledger_Name'].isin(
            company_data[entity]['names'])].index.tolist()
        interco_filt = (data['Voucher Date'] <= end_date) & (
            data['Ledger_Code'].isin(interco_ids))
        interco_df: pd.DataFrame = data.loc[interco_filt, ['Amount']]
        interco_df['Description'] = entity
        interco_df = interco_df.groupby(
            by=['Description'], as_index=False).sum()
        interco_final = pd.concat([interco_final, interco_df])
    interco_final = interco_final.loc[interco_final['Amount'] != 0]
    rpr: float = interco_final.loc[interco_final['Amount'] < 0, 'Amount'].sum()
    rpp: float = interco_final.loc[interco_final['Amount'] > 0, 'Amount'].sum()
    rpr_df: pd.DataFrame = interco_final.loc[interco_final['Amount'] < 0, [
        'Description', 'Amount']].sort_values(by='Amount', ascending=True)
    rpp_df: pd.DataFrame = interco_final.loc[interco_final['Amount'] > 0, [
        'Description', 'Amount']].sort_values(by='Amount', ascending=False)
    return {'rpr': rpr, 'rpp': rpp, 'rpr_df': rpr_df, 'rpp_df': rpp_df}


def balancesheet(data: pd.DataFrame, end_date: datetime) -> pd.DataFrame:
    offset_accounts: list = company_info[0]['data']['offset_accounts']
    # Sum total of offset_accounts is zero. i.e. PDC
    interco_acc_names: list = [i for j in [
        company_data[i]['names'] for i in company_data] for i in j]
    interco_acc_codes: list = dCoAAdler.loc[dCoAAdler['Ledger_Name'].isin(
        interco_acc_names)].index.tolist()

    exclude_bs_codes: list = offset_accounts + interco_acc_codes

    bs_filt = (data['Voucher Date'] <= end_date) & (~data['Ledger_Code'].isin(exclude_bs_codes)) & (
        data['Fourth_Level_Group_Name'].isin(['Assets', 'Liabilities', 'Equity']))
    is_filt = (data['Voucher Date'] <= end_date) & (
        data['Fourth_Level_Group_Name'].isin(['Income', 'Expenses']))

    dr_in_ap = data.loc[
        (data['Second_Level_Group_Name'] == 'Accounts Payables') & (data['Voucher Date'] <= end_date), ['Ledger_Code',
                                                                                                        'Amount']].groupby(
        by='Ledger_Code').sum()
    # returns negative figure
    dr_in_ap = dr_in_ap.loc[dr_in_ap['Amount'] < 0, 'Amount'].sum()
    cr_in_ar = data.loc[
        (data['Second_Level_Group_Name'] == 'Trade Receivables') & (data['Voucher Date'] <= end_date), ['Ledger_Code',
                                                                                                        'Amount']].groupby(
        by='Ledger_Code').sum()
    # returns positive figure
    cr_in_ar = cr_in_ar.loc[cr_in_ar['Amount'] > 0, 'Amount'].sum()

    bs_data: pd.DataFrame = data.loc[bs_filt, ['Second_Level_Group_Name', 'Amount']].groupby(
        by=['Second_Level_Group_Name'], as_index=False).sum().rename(
        columns={'Second_Level_Group_Name': 'Description'}).set_index(keys='Description')
    cum_profit: float = data.loc[is_filt, 'Amount'].sum()
    rounding_diff: float = data.loc[data['Voucher Date'] <= end_date, 'Amount'].sum()
    interco: dict = interco_bal(data=merged, end_date=end_date)
    rpr: float = interco.get('rpr')
    rpp: float = interco.get('rpp')
    rpr_row: pd.DataFrame = pd.DataFrame(data={'Amount': [rpr]}, index=[
        'Due From Related Parties'])
    rpp_row: pd.DataFrame = pd.DataFrame(
        data={'Amount': [rpp]}, index=['Due To Related Parties'])
    bs_data.loc['Accounts Payables',
    'Amount'] = bs_data.loc['Accounts Payables', 'Amount'] - dr_in_ap
    bs_data.loc['Other Receivable',
    'Amount'] = bs_data.loc['Other Receivable', 'Amount'] + dr_in_ap
    bs_data.loc['Trade Receivables',
    'Amount'] = bs_data.loc['Trade Receivables', 'Amount'] + cr_in_ar
    bs_data.loc['Accruals & Other Payables', 'Amount'] = bs_data.loc[
                                                             'Accruals & Other Payables', 'Amount'] - cr_in_ar - rounding_diff
    bs_data.loc['Retained Earnings',
    'Amount'] = bs_data.loc['Retained Earnings', 'Amount'] + cum_profit
    bs_data = pd.concat([bs_data, rpr_row, rpp_row], ignore_index=False)

    ca: float = bs_data.loc['Cash & Cash Equivalents', 'Amount'] + bs_data.loc['Inventory', 'Amount'] + bs_data.loc[
        'Other Receivable', 'Amount'] + bs_data.loc['Trade Receivables', 'Amount'] + bs_data.loc[
                    'Due From Related Parties', 'Amount']
    nca: float = (bs_data.loc['Intangible Assets', 'Amount'] if 'Intangible Assets' in bs_data.index else 0) + \
                 bs_data.loc[
                     'Property, Plant  & Equipment', 'Amount'] + \
                 bs_data.loc[
                     'Right of use Asset', 'Amount']
    equity: float = bs_data.loc['Retained Earnings', 'Amount'] + bs_data.loc['Share Capital', 'Amount'] + (bs_data.loc[
                                                                                                               'Statutory Reserves', 'Amount'] if 'Statutory Reserves' in bs_data.index else 0)
    cl: float = bs_data.loc['Accounts Payables', 'Amount'] + bs_data.loc['Accruals & Other Payables', 'Amount'] + \
                bs_data.loc[
                    'Due To Related Parties', 'Amount']
    ncl: float = bs_data.loc['Provisions', 'Amount'] + bs_data.loc[
        'Lease Liabilities', 'Amount'] if 'Lease Liabilities' in bs_data.index else 0

    ta: float = ca + nca
    tl: float = cl + ncl
    tle: float = tl + equity

    cl_row: pd.DataFrame = pd.DataFrame(
        data={'Amount': [cl]}, index=['Current Liabilities'])
    ncl_row: pd.DataFrame = pd.DataFrame(data={'Amount': [ncl]}, index=[
        'Non-Current Liabilities'])
    tl_row: pd.DataFrame = pd.DataFrame(
        data={'Amount': [tl]}, index=['Total Liabilities'])
    ca_row: pd.DataFrame = pd.DataFrame(
        data={'Amount': [ca]}, index=['Current Assets'])
    nca_row: pd.DataFrame = pd.DataFrame(
        data={'Amount': [nca]}, index=['Non-Current Assets'])
    ta_row: pd.DataFrame = pd.DataFrame(
        data={'Amount': [ta]}, index=['Total Assets'])
    equity_row: pd.DataFrame = pd.DataFrame(
        data={'Amount': [equity]}, index=['Total Equity'])
    tle_row: pd.DataFrame = pd.DataFrame(data={'Amount': [tle]}, index=[
        'Total Equity & Liabilities'])

    bs_data = pd.concat([bs_data, cl_row, ncl_row, tl_row, equity_row, ca_row, nca_row, ta_row, tle_row],
                        ignore_index=False)

    bs_data = bs_data.loc[bs_data['Amount'] != 0]
    bs_data['Amount'] = bs_data['Amount'] * -1
    return bs_data


financial_periods_bs: list = sorted(list(
    set([end_date, datetime(year=end_date.year - 1, month=end_date.month, day=end_date.day)] + list(
        pd.date_range(start=fGL['Voucher Date'].min(), end=end_date, freq='Y')))), reverse=True)
bscombined: pd.DataFrame = pd.DataFrame()
for f_year in financial_periods_bs:
    bs: pd.DataFrame = balancesheet(data=merged, end_date=f_year).rename(columns={'Amount': f'{f_year.date()}'})
    bscombined = pd.concat([bscombined, bs], axis=1)
bscombined = bscombined.reset_index().rename(columns={'index': 'Description'})

financial_periods_pl: list = sorted(list(
    set([end_date] + pd.date_range(start=fGL['Voucher Date'].min(), end=end_date, freq='Y').to_pydatetime().tolist())),
    reverse=True)
plcombined: pd.DataFrame = pd.DataFrame()
for f_year in financial_periods_pl:
    pl: dict = profitandloss(data=merged, end_date=f_year, start_date=datetime(year=f_year.year, month=1, day=1),
                             basic_pl=True)
    pl_period: pd.DataFrame = pl['df_basic']['cy_ytd_basic'].rename(columns={'Amount': f'{f_year.date()}'}).set_index(
        keys='Description')
    plcombined = pd.concat([plcombined, pl_period], axis=1)
plcombined = plcombined.reset_index()


def bsratios() -> dict:
    # current_ratio
    # gearing_ratio
    # dso
    # dpo
    # ccc
    pass


def plratios(df_pl: pd.DataFrame) -> dict:
    plmeasures: dict = {
        'gp': {'cy_cp_basic': 0, 'cy_ytd_basic': 0, 'cy_pp_basic': 0, 'py_cp_basic': 0, 'py_ytd_basic': 0,
               'cy_cp_basic_bud': 0, 'cy_ytd_basic_bud': 0},
        'np': {'cy_cp_basic': 0, 'cy_ytd_basic': 0, 'cy_pp_basic': 0, 'py_cp_basic': 0, 'py_ytd_basic': 0,
               'cy_cp_basic_bud': 0, 'cy_ytd_basic_bud': 0},
        'ebitda': {'cy_cp_basic': 0, 'cy_ytd_basic': 0, 'cy_pp_basic': 0, 'py_cp_basic': 0, 'py_ytd_basic': 0,
                   'cy_cp_basic_bud': 0, 'cy_ytd_basic_bud': 0}}

    for measure in plmeasures.keys():
        for k, v in df_pl['df_basic'].items():
            if k == 'cy_ytd_basic_monthwise':
                continue
            else:
                df: pd.DataFrame = v.set_index('Description')
                if measure == 'gp':
                    ratio: float = df.loc['Gross Profit', 'Amount'] / df.loc['Total Revenue', 'Amount'] * 100
                if measure == 'np':
                    ratio: float = df.loc['Net Profit', 'Amount'] / df.loc['Total Revenue', 'Amount'] * 100
                if measure == 'ebitda':
                    ratio: float = df.loc['Net Profit', 'Amount'] / df.loc['Total Revenue', 'Amount'] * 100
                plmeasures[measure][k] = ratio
    return plmeasures


def settlement_days(invoices: list) -> int:
    col_days: list = []
    for invoice in invoices:
        inv_value: float = fCollection.loc[(fCollection['invoice_number'] == invoice), 'invoice_amount'].head(1)
        total_collection: float = fCollection.loc[(fCollection['invoice_number'] == invoice) & (
                fCollection['voucher_date'] <= end_date), 'voucher_amount'].sum()
        if (inv_value - total_collection) == 0:
            last_date: datetime = fCollection.loc[(fCollection['invoice_number'] == invoice) & (
                    fCollection['voucher_date'] <= end_date), 'voucher_date'].max()
            inv_date: datetime = fCollection.loc[(fCollection['invoice_number'] == invoice), 'invoice_date'].head()
            col_days.append(last_date - inv_date)
    return statistics.median(col_days)


def cust_ageing(customer: str) -> pd.DataFrame:
    ledgers: list = dCustomers.loc[(dCustomers['Cus_Name'] == customer), 'Ledger_Code'].tolist()
    credit_days: int = dCustomers.loc[dCustomers['Cus_Name'].isin([customer]), 'Credit_Days'].head(1)
    invoices: list = list(set(fCollection.loc[fCollection['ledger_code'].isin(ledgers), 'invoice_number'].tolist()))
    cust_soa: pd.DataFrame = fCollection.loc[
        (fCollection['voucher_date'] <= end_date) & (fCollection['invoice_number'] == invoices), ['invoice_date',
                                                                                                  'invoice_amount',
                                                                                                  'voucher_amount',
                                                                                                  'invoice_number']]
    inv_value_list: list = []
    age_bracket_list: list = []
    ranges = [(0, 'Not Due'), (30, '1-30'), (60, '31-60'),
              (90, '61-90'), (120, '91-120'), (121, '121-150'),
              (151, '151-180'), (181, '181-210'), (211, '211-240'),
              (241, '241-270'), (271, '271-300'), (300, '301-330'),
              (331, '331-360'), (float('inf'), 'More than 361')]
    for invoice in invoices:
        total_collection: float = cust_soa.loc[(cust_soa['invoice_number'] == invoice), 'voucher_amount'].sum()
        inv_value: float = cust_soa.loc[(cust_soa['invoice_number'] == invoice), 'invoice_amount'].head(1)
        if (inv_value - total_collection) != 0:
            inv_value_list.append(inv_value - total_collection)
            days_passed: int = (end_date - cust_soa.loc[cust_soa['invoice_date']].head(1) - credit_days).days
            for threshold, label in ranges:
                if days_passed <= threshold:
                    age_bracket_list.append(label)
                break
    outstanding_df: pd.DataFrame = pd.DataFrame(
        data={'Inv_Amount': inv_value_list, 'Age Bracket': age_bracket_list}).groupby(by='Age Bracket').sum()
    return outstanding_df


def customer_ratios(customers: list, fInvoices: pd.DataFrame, end_date: datetime, fCollection: pd.DataFrame,
                    dCustomer: pd.DataFrame) -> dict:
    customer: str = ''

    customer_since: datetime = fInvoices.loc[(fInvoices['Cus_Name'] == customer), 'Invoice_Date'].min()
    total_revenue: float = fInvoices.loc[
        (fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date), 'Net_Amount'].sum()
    cust_invoices: list = fInvoices.loc[(fInvoices['Cus_Name'] == customer), 'Invoice_Number'].to_list()
    last_receipt_dt: datetime = fCollection.loc[
        fCollection['invoice_number'].isin(cust_invoices), 'voucher_date'].sort_values(by='voucher_date').tail(1)
    last_receipt_number: str = fCollection.loc[(fCollection['invoice_number'].isin(cust_invoices)) & (
            fCollection['voucher_date'] == last_receipt_dt), 'voucher_number'].tail(1)
    last_receipt_amt: float = fCollection.loc[
        (fCollection['voucher_number'] == last_receipt_number), 'voucher_amount'].groupby(by='voucher_number').sum()
    cy_cp_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
            fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=end_date.month,
                                                  day=1)), 'invoice_amount'].sum()
    cy_pp_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
            fInvoices['Invoice_Date'] <= end_date.replace(day=1) - timedelta(days=1)) & (
                                             fInvoices['Invoice_Date'] >= end_date + relativedelta(day=1,
                                                                                                   months=-1)), 'invoice_amount'].sum()
    cy_ytd_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
            fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=1, day=1)), 'invoice_amount'].sum()
    py_ytd_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
            fInvoices['Invoice_Date'] <= datetime(year=end_date - 1, month=end_date.month, day=end_date.date)) & (
                                              fInvoices['Invoice_Date'] >= datetime(year=end_date.year - 1, month=1,
                                                                                    day=1)), 'invoice_amount'].sum()
    py_cp_rev: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
            fInvoices['Invoice_Date'] <= datetime(year=end_date - 1, month=end_date.month, day=end_date.date)) & (
                                             fInvoices['Invoice_Date'] >= datetime(year=end_date.year - 1,
                                                                                   month=end_date.month,
                                                                                   day=1)), 'invoice_amount'].sum()
    collection_median: float = settlement_days(invoices=cust_invoices)
    credit_days: int = dCustomer.loc[dCustomers['Cus_Name'].isin([customer]), 'Credit_Days'].head(1)
    date_established: datetime = dCustomer.loc[dCustomer['Cus_Name'].isin([customer]), 'Date_Established'].head(1)
    outstanding_bal: float = fGL.loc[
        (fGL['Ledger_Code'].isin(dCustomer.loc[dCustomer['Cus_Name'].isin([customer]), 'Ledger_Code'].tolist())) & (
                fGL['Voucher Date'] <= end_date), 'Amount'].sum()
    cy_cp_rev_contrib_pct: float = fInvoices.loc[
                                       (fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
                                               fInvoices['Invoice_Date'] >= datetime(year=end_date.year,
                                                                                     month=end_date.month,
                                                                                     day=1)), 'invoice_amount'].sum() / \
                                   fInvoices.loc[(fInvoices['Invoice_Date'] <= end_date) & (
                                           fInvoices['Invoice_Date'] >= datetime(year=end_date.year,
                                                                                 month=end_date.month,
                                                                                 day=1)), 'invoice_amount'].sum() * 100
    cy_ytd_rev_contrib_pct: float = fInvoices.loc[(fInvoices['Cus_Name'] == customer) & (
            fInvoices['Invoice_Date'] <= end_date) & (fInvoices['Invoice_Date'] >= datetime(year=end_date.year,
                                                                                            month=1,
                                                                                            day=1)), 'invoice_amount'].sum() / \
                                    fInvoices.loc[(fInvoices['Invoice_Date'] <= end_date) & (
                                            fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=1,
                                                                                  day=1)), 'invoice_amount'].sum() * 100
    monthyly_rev: pd.DataFrame = fInvoices.loc[
        (fInvoices['Cus_Name'] == customer) & (fInvoices['Invoice_Date'] <= end_date) & (
                fInvoices['Invoice_Date'] >= datetime(year=end_date.year, month=1, day=1)), ['Invoice_Date',
                                                                                             'invoice_amount']].groupby(
        by=['Invoice_Date']).sum().rename(columns={'Invoice_Date': 'Month', 'invoice_amount': 'Amount'})
    ageing: pd.DataFrame = cust_ageing(customer=customer)

    stats: dict = {'customer_since': customer_since, 'total_revenue': total_revenue, 'credit_score': 0,
                   'last_receipt_amt': last_receipt_amt, 'cy_cp_rev': cy_cp_rev, 'cy_pp_rev': cy_pp_rev,
                   'last_receipt_dt': last_receipt_dt,
                   'cy_ytd_rev': cy_ytd_rev, 'py_cp_rev': py_cp_rev, 'py_ytd_rev': py_ytd_rev,
                   'collection_median': collection_median, 'credit_days': credit_days, 'last_sales_person': 0,
                   'customer_gp': 0, 'outstanding_bal': outstanding_bal, 'ageing': ageing,
                   'date_established': date_established,
                   'cy_cp_rev_contrib_pct': cy_cp_rev_contrib_pct, 'cy_ytd_rev_contrib_pct': cy_ytd_rev_contrib_pct,
                   'cy_cp_roi': 0,
                   'cy_ytd_roi': 0, 'monthyly_rev': monthyly_rev, 'remarks': 0}

    return stats


def sales_person(salesmen: str,dEmployee:pd.DataFrame) -> dict:
    emp_id :str = ''
    doj:datetime = dEmployee.loc[emp_id,'doj']

    stats: dict = {'doj':doj, 'target': 0, 'cy_cp_rev': 0,
                   'cy_ytd_rev': 0, 'cy_cp_rev_org': 0, 'cy_ytd_rev_org': 0,
                   'new_customers_added': 0, 'cy_cp_gp': 0, 'cy_ytd_gp': 0,
                   'ar_balance': 0, 'monthly_rev': 0, 'cy_cp_rev_contrib_pct': 0,
                   'cy_ytd_rev_contri_pct': 0}
    

    return stats


def revenue(end_date: datetime, data: pd.DataFrame) -> dict:
    rev_filt = (data['Third_Level_Group_Name'] == 'Direct Income') & (
            data['Voucher Date'] <= end_date)
    rev_division: pd.DataFrame = data.loc[rev_filt, ['Voucher Date', 'Amount', 'Second_Level_Group_Name']].groupby(
        by=['Voucher Date', 'Second_Level_Group_Name'], as_index=False).sum()
    sales_invoices: list = list(
        set(data.loc[rev_filt, 'Voucher Number'].tolist()))
    total_invoices: list = list(set(fInvoices['Invoice_Number'].tolist()))
    worked_invoices: list = [
        inv for inv in sales_invoices if inv in total_invoices]
    rev_category: pd.DataFrame = data.loc[
        (data['Voucher Number'].isin(worked_invoices)) & (data['Third_Level_Group_Name'] == 'Direct Income'), [
            'Voucher Number', 'Amount', 'Voucher Date']].rename(
        columns={'Voucher Number': 'Invoice_Number'})
    rev_category: pd.DataFrame = pd.merge(left=rev_category, right=fInvoices[['Invoice_Number', 'Type']],
                                          on='Invoice_Number', how='left').drop(columns=['Invoice_Number']).groupby(
        by=['Voucher Date', 'Type'], as_index=False).sum()
    return {'rev_division': rev_division, 'rev_category': rev_category}


def closing_date(row) -> datetime:
    """Add credit period (in days) to the voucher date and convert that date to end of the month

    Args:
        row (_type_): a row in dataframe

    Returns:
        datetime: last date of the month to which voucher becomes due
    """
    ledger_code: int = row['Ledger Code']
    if ledger_code in dCustomers.index:
        credit_days: int = int(dCustomers.loc[ledger_code, 'Credit_Days'])
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
    filt_net_rev = (fGL['Voucher Date'] >= start_date) & (fGL['Voucher Date'] <= end_date) & (
        fGL['Transaction Type'].isin(VOUCHER_TYPES)) & (fGL['Fourth Level Group Name'] == 'Assets')
    fGL = fGL.loc[filt_net_rev]
    fGL['Due Date'] = fGL.apply(closing_date, axis=1)
    df_already_collected: pd.DataFrame = fGL
    start_date: datetime = row['Due Date'].replace(day=1)
    period_filt = (df_already_collected['Due Date'] >= start_date) & (
            df_already_collected['Due Date'] <= row['Due Date'])
    due_inv_list: list = list(
        set(df_already_collected.loc[period_filt, 'Voucher Number'].tolist()))
    already_collected_receipts: pd.DataFrame = receipts_recorded(fGL=fGL, fCollection=fCollection)
    collected_filt = (already_collected_receipts['Invoice_number'].isin(due_inv_list)) & (
            already_collected_receipts['Voucher_Date'] < start_date)
    amount: float = already_collected_receipts.loc[collected_filt, 'Credit'].sum(
    )
    return amount


def collection() -> pd.DataFrame:
    receipts: pd.DataFrame = receipts_recorded(
        fGL=fGL, fCollection=fCollection)
    already_collected_receipts: pd.DataFrame = receipts

    # filters the collection date based on the selection
    filt_collection = (receipts['Voucher_Date'] >= start_date) & (
            receipts['Voucher_Date'] <= end_date)
    receipts = receipts.loc[filt_collection]
    # convert collection date to last date of the month, so it can be grouped to know total collected per period.
    receipts['Voucher_Date'] = receipts.apply(
        lambda row: row['Voucher_Date'] + relativedelta(day=31), axis=1)
    receipts = receipts.groupby(by=['Voucher_Date'], as_index=False)[
        'Credit'].sum()
    receipts.rename(columns={'Voucher_Date': 'Due Date',
                             'Credit': 'Actual'}, inplace=True)
    # Reasons for Finance / Receipt total for a period not match with 'Actual' in this report
    # 1. Credit notes are part of 'Actual' in this report
    # 2. Receipts other than from customers i.e. Employee Receivable is not part of this report
    # 3. Receipts that were not allocated to invoices are not part of this report.
    # for 3 above check fCollection/Invoice Number Contains RV/CN and Payment Date ->Blank

    fGL = fGL.groupby(by=['Due Date'], as_index=False)[
        'Debit Amount'].sum()
    fGL['Already_Collected'] = fGL.apply(already_collected, axis=1)
    fGL['Debit Amount'] = fGL['Debit Amount'] - fGL['Already_Collected']
    fGL.drop(columns=['Already_Collected'], inplace=True)
    fGL = fGL.loc[(fGL['Due Date'] >= start_date)
                  & (fGL['Due Date'] <= end_date)]
    fGL.rename(columns={'Debit Amount': 'Target'}, inplace=True)

    combined: pd.DataFrame = pd.concat([receipts.set_index('Due Date'), fGL.set_index('Due Date')], axis=1,
                                       join='outer').reset_index()


df_pl: dict = profitandloss(basic_pl=True, data=merged, start_date=start_date, end_date=end_date, full_pl=True)
cy_cp_basic: pd.DataFrame = df_pl['df_basic']['cy_cp_basic']
cy_ytd_basic: pd.DataFrame = df_pl['df_basic']['cy_ytd_basic']
cy_pp_basic: pd.DataFrame = df_pl['df_basic']['cy_pp_basic']
py_cp_basic: pd.DataFrame = df_pl['df_basic']['py_cp_basic']
py_ytd_basic: pd.DataFrame = df_pl['df_basic']['py_ytd_basic']
cy_cp_basic_bud: pd.DataFrame = df_pl['df_basic']['cy_cp_basic_bud']
cy_ytd_basic_bud: pd.DataFrame = df_pl['df_basic']['cy_ytd_basic_bud']

cp_month: pd.DataFrame = pd.concat(
    [cy_cp_basic.set_index('Description'), cy_pp_basic.set_index('Description'), py_cp_basic.set_index('Description'),
     cy_cp_basic_bud.set_index('Description')],
    axis=1, join='outer').reset_index()

document = Document()
doc = first_page(document=document, report_date=end_date)
document.add_page_break()
tbl_month_basic = document.add_table(rows=1, cols=5)
heading_cells = tbl_month_basic.rows[0].cells
heading_cells[0].text = 'Description'
heading_cells[1].text = 'Current Month'
heading_cells[2].text = 'Previous Month'
heading_cells[3].text = 'SPLY'
heading_cells[4].text = 'Budget'

for _, row in cp_month.iterrows():
    cells = tbl_month_basic.add_row().cells
    cells[0].text = str(row['Description'])
    cells[1].text = f"{row.iloc[1]:,.0f}" if row.iloc[1] >= 0 else f"({abs(row.iloc[1]):,.0f})"
    cells[2].text = f"{row.iloc[2]:,.0f}" if row.iloc[2] >= 0 else f"({abs(row.iloc[2]):,.0f})"
    cells[3].text = f"{row.iloc[3]:,.0f}" if row.iloc[3] >= 0 else f"({abs(row.iloc[3]):,.0f})"
    cells[4].text = f"{row.iloc[4]:,.0f}" if row.iloc[4] >= 0 else f"({abs(row.iloc[4]):,.0f})"

tbl_month_basic.style = 'Light Grid Accent 1'
document.add_page_break()

cy_cp_full: pd.DataFrame = df_pl['df_full']['cy_cp_full']
cy_pp_full: pd.DataFrame = df_pl['df_full']['cy_pp_full']
py_cp_full: pd.DataFrame = df_pl['df_full']['py_cp_full']
cy_cp_full_bud: pd.DataFrame = df_pl['df_full']['cy_cp_full_bud']

cp_month_full: pd.DataFrame = pd.concat(
    [cy_cp_full.set_index('Description'), cy_pp_full.set_index('Description'), py_cp_full.set_index('Description'),
     cy_cp_full_bud.set_index('Description')],
    axis=1, join='outer').reset_index()

tbl_month_full = document.add_table(rows=1, cols=5)
heading_cells = tbl_month_full.rows[0].cells
heading_cells[0].text = 'Description'
heading_cells[1].text = 'Current Month'
heading_cells[2].text = 'Previous Month'
heading_cells[3].text = 'SPLY'
heading_cells[4].text = 'Budget'

for _, row in cp_month_full.iterrows():
    cells = tbl_month_full.add_row().cells
    cells[0].text = str(row['Description'])
    cells[1].text = f"{row.iloc[1]:,.0f}" if row.iloc[1] >= 0 else f"({abs(row.iloc[1]):,.0f})"
    cells[2].text = f"{row.iloc[2]:,.0f}" if row.iloc[2] >= 0 else f"({abs(row.iloc[2]):,.0f})"
    cells[3].text = f"{row.iloc[3]:,.0f}" if row.iloc[3] >= 0 else f"({abs(row.iloc[3]):,.0f})"
    cells[4].text = f"{row.iloc[4]:,.0f}" if row.iloc[4] >= 0 else f"({abs(row.iloc[4]):,.0f})"
tbl_month_full.style = 'Light Grid Accent 2'
document.add_page_break()

cy_ytd_basic: pd.DataFrame = df_pl['df_basic']['cy_ytd_basic']
py_ytd_basic: pd.DataFrame = df_pl['df_basic']['py_ytd_basic']
cy_ytd_basic_bud: pd.DataFrame = df_pl['df_basic']['cy_ytd_basic_bud']

cp_ytd: pd.DataFrame = pd.concat(
    [cy_ytd_basic.set_index('Description'), py_ytd_basic.set_index('Description'),
     cy_ytd_basic_bud.set_index('Description')], axis=1, join='outer').reset_index()

tbl_ytd_basic = document.add_table(rows=1, cols=4)
heading_cells = tbl_ytd_basic.rows[0].cells
heading_cells[0].text = 'Description'
heading_cells[1].text = 'YTD CY'
heading_cells[2].text = 'YTD PY'
heading_cells[3].text = 'Budget'

for _, row in cp_ytd.iterrows():
    cells = tbl_ytd_basic.add_row().cells
    cells[0].text = str(row['Description'])
    cells[1].text = f"{row.iloc[1]:,.0f}" if row.iloc[1] >= 0 else f"({abs(row.iloc[1]):,.0f})"
    cells[2].text = f"{row.iloc[2]:,.0f}" if row.iloc[2] >= 0 else f"({abs(row.iloc[2]):,.0f})"
    cells[3].text = f"{row.iloc[3]:,.0f}" if row.iloc[3] >= 0 else f"({abs(row.iloc[3]):,.0f})"

tbl_ytd_basic.style = 'Light Grid Accent 3'
document.add_page_break()

cy_ytd_basic_monthwise: pd.DataFrame = df_pl['df_basic']['cy_ytd_basic_monthwise']

tbl_monthwise_basic = document.add_table(rows=1, cols=cy_ytd_basic_monthwise.shape[1])
heading_cells = tbl_monthwise_basic.rows[0].cells

for i in range(cy_ytd_basic_monthwise.shape[1]):
    if i == 0:
        heading_cells[i].text = 'Description'
    else:
        heading_cells[i].text = list(cy_ytd_basic_monthwise.columns)[i].strftime('%b')

for _, row in cy_ytd_basic_monthwise.iterrows():
    cells = tbl_monthwise_basic.add_row().cells
    for j in range(len(row)):
        if j == 0:
            cells[0].text = str(row['Description'])
        else:
            cells[j].text = f"{row.iloc[j]:,.0f}" if row.iloc[j] >= 0 else f"({abs(row.iloc[j]):,.0f})"

tbl_monthwise_basic.style = 'Light Grid Accent 5'
document.add_page_break()

df_rev: dict = revenue(end_date=end_date, data=merged)
rev_division: pd.DataFrame = df_rev['rev_division']
rev_division_plot: pd.DataFrame = rev_division.copy()


def plotting_period(end_date: datetime, months: int) -> datetime:
    start_date: datetime = datetime(year=end_date.year, month=1, day=1)
    delta = relativedelta(dt1=end_date, dt2=start_date)
    period: int = delta.months + (delta.years * 12) + 1
    if period < 6:
        start_date = end_date - relativedelta(months=months - 1)
        start_date = datetime(year=start_date.year, month=start_date.month, day=1)
    else:
        start_date
    return start_date


rev_division_line: pd.DataFrame = rev_division_plot.loc[(rev_division_plot['Voucher Date'] <= end_date) & (
        rev_division_plot['Voucher Date'] >= plotting_period(end_date=end_date, months=6))].pivot_table(
    index='Voucher Date', columns='Second_Level_Group_Name', values='Amount',
    aggfunc='sum', fill_value=0).reset_index().rename(columns={'Voucher Date': 'Period'}).set_index(keys='Period')
rev_division_pie_ytd: pd.DataFrame = rev_division_plot.loc[(rev_division_plot['Voucher Date'] <= end_date) & (
        rev_division_plot['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1)), [
    'Second_Level_Group_Name', 'Amount']].groupby(by='Second_Level_Group_Name').sum().reset_index().rename(
    columns={'Second_Level_Group_Name': 'Category'}).set_index(keys='Category')
rev_division_pie_month: pd.DataFrame = rev_division_plot.loc[(rev_division_plot['Voucher Date'] <= end_date) & (
        rev_division_plot['Voucher Date'] >= datetime(year=end_date.year, month=end_date.month, day=1)), [
    'Second_Level_Group_Name', 'Amount']].rename(columns={'Second_Level_Group_Name': 'Category'}).set_index(
    keys='Category')

rev_div_buf = BytesIO()

rev_div, (manpower, project) = plt.subplots(nrows=2, ncols=1)
plt.style.use('ggplot')
manpower.plot(rev_division_line.index, rev_division_line['Manpower'], label='Manpower')
project.plot(rev_division_line.index, rev_division_line['Projects'], label='Projects')
manpower.set_title('Manpower Division Monthly Revenue')
project.set_title('ELV Division Monthly Revenue')
manpower.set_xticks(ticks=rev_division_line.index, labels=[i.strftime('%b') for i in rev_division_line.index])
project.set_xticks(ticks=rev_division_line.index, labels=[i.strftime('%b') for i in rev_division_line.index])

current_values_project = project.get_yticks()
project.set_yticklabels(['{:,}'.format(int(i)) for i in current_values_project])

current_values_project = manpower.get_yticks()
manpower.set_yticklabels(['{:,}'.format(int(i)) for i in current_values_project])

manpower.grid()
project.grid()
plt.tight_layout()
plt.savefig(rev_div_buf, format='png')
plt.close(rev_div)

rev_div_total_buf = BytesIO()
div_total, (month_div_pie, ytd_div_pie) = plt.subplots(nrows=2, ncols=1)
plt.style.use('ggplot')
month_div_pie.set_title('Month')
month_div_pie.pie(x=rev_division_pie_month['Amount'], labels=rev_division_pie_month.index, autopct='%.1f%%')
ytd_div_pie.set_title('YTD')
ytd_div_pie.pie(x=rev_division_pie_ytd['Amount'], labels=rev_division_pie_ytd.index, autopct='%.1f%%')
plt.savefig(rev_div_total_buf, format='png')
plt.close(div_total)

rev_division = rev_division.loc[(rev_division['Voucher Date'] <= end_date) & (
        rev_division['Voucher Date'] >= plotting_period(end_date=end_date, months=6))].pivot_table(
    index='Second_Level_Group_Name', columns='Voucher Date', values='Amount',
    aggfunc='sum', fill_value=0).reset_index().rename(columns={'Second_Level_Group_Name': 'Description'})

tbl_monthwise_rev_div = document.add_table(rows=1, cols=rev_division.shape[1])
heading_cells = tbl_monthwise_rev_div.rows[0].cells

for i in range(rev_division.shape[1]):
    if i == 0:
        heading_cells[i].text = 'Description'
    else:
        heading_cells[i].text = list(rev_division.columns)[i].strftime('%b')

for _, row in rev_division.iterrows():
    cells = tbl_monthwise_rev_div.add_row().cells
    for j in range(len(row)):
        if j == 0:
            cells[0].text = str(row['Description'])
        else:
            cells[j].text = f"{row.iloc[j]:,.0f}" if row.iloc[j] >= 0 else f"({abs(row.iloc[j]):,.0f})"

tbl_monthwise_rev_div.style = 'Light Grid Accent 6'
document.add_page_break()

rev_category: pd.DataFrame = df_rev['rev_category']
rev_category = rev_category.loc[(rev_category['Voucher Date'] <= end_date) & (
        rev_category['Voucher Date'] >= plotting_period(end_date=end_date, months=6))].pivot_table(index='Type',
                                                                                                   columns='Voucher '
                                                                                                           'Date',
                                                                                                   values='Amount',
                                                                                                   aggfunc='sum',
                                                                                                   fill_value=0).reset_index().rename(
    columns={'Type': 'Description'})
rev_category_pie: pd.DataFrame = df_rev['rev_category']
rev_category_pie_ytd: pd.DataFrame = rev_category_pie.loc[(rev_category_pie['Voucher Date'] <= end_date) & (
        rev_category_pie['Voucher Date'] >= datetime(year=end_date.year, month=1, day=1)), ['Type',
                                                                                            'Amount']].groupby(
    by='Type').sum()
rev_category_pie_month: pd.DataFrame = rev_category_pie.loc[(rev_category_pie['Voucher Date'] <= end_date) & (
        rev_category_pie['Voucher Date'] >= datetime(year=end_date.year, month=end_date.month, day=1)), ['Type',
                                                                                                         'Amount']].groupby(
    by='Type').sum()

rev_cat_total_buf = BytesIO()
cat_total, (month_cat_pie, ytd_cat_pie) = plt.subplots(nrows=2, ncols=1)
plt.style.use('ggplot')
month_cat_pie.set_title('Month')
month_cat_pie.pie(x=rev_category_pie_month['Amount'], labels=rev_category_pie_month.index, autopct='%.0f%%')
ytd_cat_pie.set_title('YTD')
ytd_cat_pie.pie(x=rev_category_pie_ytd['Amount'], labels=rev_category_pie_ytd.index, autopct='%.0f%%')
plt.savefig(rev_cat_total_buf, format='png')
plt.close(cat_total)

tbl_monthwise_rev_cat = document.add_table(rows=1, cols=rev_category.shape[1])
heading_cells = tbl_monthwise_rev_cat.rows[0].cells

for i in range(rev_category.shape[1]):
    if i == 0:
        heading_cells[i].text = 'Description'
    else:
        heading_cells[i].text = list(rev_category.columns)[i].strftime('%b')

for _, row in rev_category.iterrows():
    cells = tbl_monthwise_rev_cat.add_row().cells
    for j in range(len(row)):
        if j == 0:
            cells[0].text = str(row['Description'])
        else:
            cells[j].text = f"{row.iloc[j]:,.0f}" if row.iloc[j] >= 0 else f"({abs(row.iloc[j]):,.0f})"

tbl_monthwise_rev_cat.style = 'Light Grid'
document.add_page_break()
rev_div_buf.seek(0)
doc.add_picture(rev_div_buf)
document.add_page_break()
rev_div_total_buf.seek(0)
doc.add_picture(rev_div_total_buf)
document.add_page_break()
rev_cat_total_buf.seek(0)
doc.add_picture(rev_cat_total_buf)
document.add_page_break()

tbl_yearly_bs = document.add_table(rows=1, cols=bscombined.shape[1])
heading_cells = tbl_yearly_bs.rows[0].cells
for i in range(bscombined.shape[1]):
    if i == 0:
        heading_cells[i].text = 'Description'
    else:
        heading_cells[i].text = list(bscombined.columns)[i]

for _, row in bscombined.iterrows():
    cells = tbl_yearly_bs.add_row().cells
    for j in range(len(row)):
        if j == 0:
            cells[0].text = str(row['Description'])
        else:
            cells[j].text = f"{row.iloc[j]:,.0f}" if row.iloc[j] >= 0 else f"({abs(row.iloc[j]):,.0f})"

tbl_yearly_bs.style = 'Light Grid Accent 5'
document.add_page_break()

tbl_yearly_pl = document.add_table(rows=1, cols=plcombined.shape[1])
heading_cells = tbl_yearly_pl.rows[0].cells
for i in range(plcombined.shape[1]):
    if i == 0:
        heading_cells[i].text = 'Description'
    else:
        heading_cells[i].text = list(plcombined.columns)[i]

for _, row in plcombined.iterrows():
    cells = tbl_yearly_pl.add_row().cells
    for j in range(len(row)):
        if j == 0:
            cells[0].text = str(row['Description'])
        else:
            cells[j].text = f"{row.iloc[j]:,.0f}" if row.iloc[j] >= 0 else f"({abs(row.iloc[j]):,.0f})"

tbl_yearly_pl.style = 'Light Grid Accent 5'
document.add_page_break()

interco: dict = interco_bal(data=merged, end_date=end_date)
rpr_df: pd.DataFrame = interco.get('rpr_df')
rpr_total_row: pd.DataFrame = pd.DataFrame(data={'Amount': [rpr_df['Amount'].sum()], 'Description': 'Total'}, index=[
    '9999'])
rpr_df = pd.concat([rpr_df, rpr_total_row], ignore_index=False)

tbl_rpr = document.add_table(rows=1, cols=2)
heading_cells = tbl_rpr.rows[0].cells
heading_cells[0].text = 'Description'
heading_cells[1].text = 'Amount'

for _, row in rpr_df.iterrows():
    cells = tbl_rpr.add_row().cells
    cells[0].text = str(row['Description'])
    cells[1].text = f"{row.iloc[1]:,.0f}" if row.iloc[1] >= 0 else f"{abs(row.iloc[1]):,.0f}"

tbl_rpr.style = 'Light Grid Accent 3'
document.add_page_break()

rpp_df: float = interco.get('rpp_df')
rpp_total_row: pd.DataFrame = pd.DataFrame(data={'Amount': [rpp_df['Amount'].sum()], 'Description': 'Total'}, index=[
    '9999'])
rpp_df = pd.concat([rpp_df, rpp_total_row], ignore_index=False)

tbl_rpp = document.add_table(rows=1, cols=2)
heading_cells = tbl_rpp.rows[0].cells
heading_cells[0].text = 'Description'
heading_cells[1].text = 'Amount'

for _, row in rpp_df.iterrows():
    cells = tbl_rpp.add_row().cells
    cells[0].text = str(row['Description'])
    cells[1].text = f"{row.iloc[1]:,.0f}" if row.iloc[1] >= 0 else f"{abs(row.iloc[1]):,.0f}"

tbl_rpp.style = 'Light Grid Accent 3'
document.add_page_break()

document.core_properties.author = "Nadun Jayathunga"
document.core_properties.keywords = ("Chief Accountant\nNasser Bin Nawaf and Partners Holdings "
                                     "W.L.L\nE-mail\tnjayathunga@nbn.qa\nTelephone\t+974 4403 0407")

doc.save('Monthly FS.docx')
convert('Monthly FS.docx')
os.unlink('Monthly FS.docx')
