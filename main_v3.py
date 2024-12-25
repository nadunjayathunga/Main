from docx import Document


from helper_3 import (welcome_page, data_sources, preprocessing, first_page, closing, page_separator, header,
                    pl_month_brief, pl_month_detailed, log_div_profit, change_orientation, pl_ytd_brief,
                    gpnpebitda_graph, plmonthwise, excpdetails, guardingbumonthwise, elvbumonthwise, credits,
                    plhistorical, bshistorical, rpbalances, apbalances, main_bs_ratios, revenue,
                    revenue_dashboard, data_output, revenue_dashboard_two,  customer_specifics,
                    revenue_movement, hrrelated, opsrelated,cohart,re_related,occupancy_report,re_rev_recon,console_hrrelated,
                    rpt_graphs,consolidated_pandl)

welcome_info: dict = welcome_page()
raw_data: dict = data_sources(engine=welcome_info['engine'], database=welcome_info['database'])
refined_data: dict = preprocessing(data=raw_data, database=welcome_info['database'])
output_data: dict = data_output(refined=refined_data, welcome_info=welcome_info)
CUST_LOGO_PATH = r'C:\Masters\images\customer'
document = Document()

first_page(document=document, report_date=welcome_info['end_date'], abbr=welcome_info['abbr'],
           long_name=welcome_info['long_name'])
document.add_page_break()
page_separator(head='Finance', document=document)
header(title='Profit & Loss for the current period', company=welcome_info['long_name'], document=document)
pl_month_brief(document=document, special=['Total Revenue', 'Gross Profit', 'Total Overhead', 'Net Profit'],
               data=output_data['cp_month'])
header(title='Complete Profit & Loss for the current period', company=welcome_info['long_name'], document=document)
pl_month_detailed(document=document, special=['Total Revenue', 'Gross Profit', 'Total Overhead', 'Net Profit'],
                  data=output_data['cp_month_full'])
header(title='Profit & Loss for Year to Date', company=welcome_info['long_name'], document=document)
pl_ytd_brief(document=document, data=output_data['cp_ytd'],
             special=['Total Revenue', 'Gross Profit', 'Total Overhead', 'Net Profit'])
document.add_page_break()
gpnpebitda_graph(document=document, end_date=welcome_info['end_date'], ratios=output_data['ratios_pandl'])
document.add_page_break()
change_orientation(document=document, method='l')
header(title='Profit & Loss for Year to Date Month-Wise', company=welcome_info['long_name'], document=document)
plmonthwise(document=document, data=output_data['cy_ytd_basic_monthwise'],
            special=['Total Revenue', 'Gross Profit', 'Total Overhead', 'Net Profit'])
document.add_page_break()
change_orientation(document=document, method='p')
excpdetails(document=document, data=output_data['merged'], end_date=welcome_info['end_date'],long_name=welcome_info['long_name'])
if welcome_info['database'] == 'nbn_logistics':
    change_orientation(document=document, method='l')
    header(title='Division wise Profitability', company=welcome_info['long_name'], document=document)
    log_div_profit(profit=output_data['cat_profit'], document=document)
    change_orientation(document=document, method='p')
elif welcome_info['database'] == 'elite_security':
    change_orientation(document=document, method='l')
    header(title='Profit & Loss for Guarding Division', company=welcome_info['long_name'], document=document)
    guardingbumonthwise(document=document, end_date=welcome_info['end_date'],
                        special=['Total Revenue', 'Gross Profit', 'Total Overhead', 'Net Profit'],
                        fBudget=refined_data['fBudget'], merged=output_data['merged'],
                        sort_order=output_data['sort_order'])
    document.add_page_break()
    header(title='Profit & Loss for ELV Division', company=welcome_info['long_name'], document=document)
    elvbumonthwise(document=document, end_date=welcome_info['end_date'],
                   special=['Total Revenue', 'Gross Profit', 'Total Overhead', 'Net Profit'],
                   fBudget=refined_data['fBudget'], merged=output_data['merged'], sort_order=output_data['sort_order'])
    document.add_page_break()
    change_orientation(document=document, method='p')
else:
    pass
header(title='Historical Profit and Loss Comparison', company=welcome_info['long_name'], document=document)
plhistorical(document=document, data=output_data['plcombined'],
             special=['Total Revenue', 'Gross Profit', 'Total Overhead', 'Net Profit'],
             sort_order=output_data['sort_order'])
document.add_page_break()
change_orientation(document=document, method='l')
header(title='Statement of Financial Position (Balance Sheet)', company=welcome_info['long_name'], document=document)
bshistorical(document=document, data=output_data['bscombined'],
             special=['Current Liabilities', 'Non Current Liabilities', 'Liabilities', 'Equity', 'Current Assets',
                      'Non Current Assets', 'Assets', 'Total Equity & Liabilities'],
             sort_order=output_data['sort_order'])
change_orientation(document=document, method='p')
header(title='Break-up of Related-Party Balances', company=welcome_info['long_name'], document=document)
rpbalances(document=document, end_date=welcome_info['end_date'], data=output_data['merged'],
           dCoAAdler=refined_data['dCoAAdler'])
if welcome_info['database'] == 'nbn_holding':
    document.add_page_break()
    rpt_graphs(end_date=welcome_info['end_date'],document=document)
document.add_page_break()
header(title='Accounts Payable Break-up', company=welcome_info['long_name'], document=document)
apbalances(document=document, fAP=refined_data['fAP'])
document.add_page_break()
main_bs_ratios(document=document, end_date=welcome_info['end_date'], bsdata=output_data['bscombined'],
               pldata=output_data['plcombined'], periods=output_data['financial_periods_bs'],database = welcome_info['database'])
document.add_page_break()
if welcome_info['database'] in ['elite_security','nbn_logistics','premium']:
    page_separator(head='Sales', document=document)
    change_orientation(document=document, method='l')
    df_rev = revenue(end_date=welcome_info['end_date'], data=output_data['merged'], 
                    fInvoices=refined_data['fInvoices'],database=welcome_info['database'],
                    fData=refined_data.get('fData'),dJobs=refined_data.get('dJobs'),dCustomer=refined_data['dCustomer'])
    revenue_dashboard(document=document, end_date=welcome_info['end_date'], months=6, database=welcome_info['database'],
                    df_rev=df_rev)
    document.add_page_break()
    revenue_dashboard_two(df_rev=df_rev, document=document, refined_data=refined_data, welcome_info=welcome_info)
    document.add_page_break()
    change_orientation(document=document, method='p')
    header(title='Key data about customers', company=welcome_info['long_name'], document=document)
    profitability: dict = customer_specifics(document=document, fInvoices=refined_data['fInvoices'],
                                            end_date=welcome_info['end_date'],
                                            dCustomer=refined_data['dCustomer'], path=CUST_LOGO_PATH,
                                            fGL=output_data['merged'],
                                            dEmployee=refined_data['dEmployee'], dExclude=refined_data.get('dExclude'),
                                            dJobs=refined_data['dJobs'],
                                            fCollection=refined_data['fCollection'], fOT=refined_data.get('fOT'),
                                            fTimesheet=refined_data.get('ftimesheet'),
                                            database=welcome_info['database'],
                                            fData=refined_data.get('fData'),fLogInv=refined_data.get('fLogInv'),fMI = refined_data.get('fMI'))
    document.add_page_break()
    revenue_movement(document=document, fInvoices=refined_data['fInvoices'], end_date=welcome_info['end_date'],database=welcome_info['database'])
if welcome_info['database'] == 'nbn_logistics':
    cohart(document=document,end_date=welcome_info['end_date'],fInvoices=refined_data['fInvoices'],fSalesTill2020=refined_data['fSalesTill2020'],dCustomer=refined_data['dCustomer'])
if welcome_info['database'] in ['elite_security','nbn_logistics','premium','nbn_holding']:
    page_separator(head='HR', document=document)
    hrrelated(document=document, database=welcome_info['database'], dEmployee=refined_data['dEmployee'],
            end_date=welcome_info['end_date'])
if welcome_info['database']=='nbn_holding':
    console_hrrelated(database=welcome_info['database'],document=document,end_date=welcome_info['end_date'])
if welcome_info['database'] == 'elite_security':
    page_separator(head='Operations', document=document)
    opsrelated(financial=output_data['cy_ytd_basic_monthwise'], document=document, dExclude=refined_data.get('dExclude'),
            end_date=welcome_info['end_date'], fTimesheet=refined_data.get('ftimesheet'),
            profitability=profitability)
if welcome_info['database'] == 'nbn_realestate':
    header(title='Unit Occupancy', company=welcome_info['long_name'], document=document)
    re_reports:dict = occupancy_report(end_date=welcome_info['end_date'],dJobs=refined_data.get('dJobs'),fInvoices=refined_data['fInvoices'],dRoom=refined_data.get('dRoom'))
    re_related(document=document,re_reports=re_reports)
    document.add_page_break()
    header(title='Revenue Reconcilliation', company=welcome_info['long_name'], document=document)
    re_rev_recon(document=document,re_reports=re_reports,fGL=output_data['merged'],end_date=welcome_info['end_date'])
    document.add_page_break()
credits(document=document,abbr=welcome_info['abbr'])
closing(document=document, abbr=welcome_info['abbr'],end_date=welcome_info['end_date'])
consolidated_pandl(welcome_info=welcome_info)

