import pandas as pd
from colorama import Fore, init
from data import company_info
import numpy as np

init()

print('\n')
for idx, company in enumerate(company_info):
    print(f'{idx}\t{company["data"]["long_name"]}')
print('\n')
company_id: int = int(input('Please enter company ID>> '))
print('\n')

file_name: str = company_info[company_id]['data']['file_name']
file_path = f'C:\Masters\{file_name}.xlsx'

"""sheets_to_check dictionary contains the sheets to be checked(As key), the coloumns to be read from each sheet mentioned as key, 
whether any test to be performed and if any test is required which column to be checked in column in that sheet.
each sheet which are considered for this testinig purpose may contain primary key (i.e emp_id in dEmployee) and one or more foreign keys (i.e emp_id in dContracts). 
foreign keys in a sheet required be tested for the existance of the same as a primary key in master list(i.e emp_id, part_number etc). 

If a new sheet has been added to a workbook follow the steps below.
1. Add the sheet name as key in sheets_to_check dictionary.
2. Add the 'cid' (company identification number) to the list 
3. Add 'usecols' to the list of columns to be read from the sheet.
4. If any test is required set 'test' to True and add the column to be checked in 'check' list.
"""
sheets_to_check: dict = {'fBudget': {'cid': ['3', '1', '7', '9', '4', '6', '10', '5', '8', '2',], 'usecols': ['ledger_code',], 'test': True, 'check': ['ledger_code',]},
                         'dCoAAdler': {'cid': ['3', '1', '7', '9', '4', '6', '10', '5', '8', '2',], 'usecols': ['ledger_code'], 'test': False},
                         'fData': {'cid': ['3'], 'usecols': ['ledger_code', 'order_id', 'service_element_code'], 'test': True, 'check': ['ledger_code', 'order_id', 'service_element_code']},
                         'dLogContract': {'cid': ['3',], 'usecols': ['order_id', 'customer_code', 'emp_id'], 'test': True, 'check': ['customer_code', 'emp_id']},
                         'dEmployee': {'cid': ['3', '1', '7', '9', '4', '6', '10', '5', '8', '2',], 'usecols': ['emp_id'], 'test': False},
                         'dCustomer': {'cid': ['3', '1', '4', '6', '8', '2','5'], 'usecols': ['customer_code'], 'test': False},
                         'dServiceTypes': {'cid': ['3',], 'usecols': ['service_element_code'], 'test': False},
                         'fGL': {'cid': ['3', '1', '7', '9', '4', '6', '10', '5', '8', '2',], 'usecols': ['ledger_code'], 'test': True, 'check': ['ledger_code']},
                         'fNotInvoiced': {'cid': ['3',], 'usecols': ['order_id', 'ledger_code'], 'test': True, 'check': ['order_id', 'ledger_code']},
                         'fCollection': {'cid': ['3', '1', '4', '6', '5', '8', '2'], 'usecols': ['ledger_code'], 'test': True, 'check': ['ledger_code']},
                         'fLogInv': {'cid': ['3',], 'usecols': ['order_id', 'customer_code', 'emp_id'], 'test': True, 'check': ['order_id', 'customer_code', 'emp_id']},
                         'dRentContract': {'cid': ['4'], 'usecols': ['order_id', 'customer_code', 'room_id'], 'test': False},
                         'dStock': {'cid': ['4', '2', '1'], 'usecols': ['part_number'], 'test': False},
                         'dRoom': {'cid': ['4'], 'usecols': ['room_id'], 'test': False},
                         'fAP': {'cid': ['4', '6', '5', '8', '2',], 'usecols': ['ledger_code'], 'test': False},
                         'dCusOrder': {'cid': ['6', '8', '5', '1'], 'usecols': ['order_id', 'customer_code', 'emp_id'], 'test': True, 'check': ['customer_code', 'emp_id']},
                         'dContract': {'cid': ['1', '2',], 'usecols': ['order_id', 'customer_code', 'emp_id'], 'test': True, 'check': ['customer_code', 'emp_id']},
                         'dOrderAMC': {'cid': ['1'], 'usecols': ['customer_code', 'emp_id', 'order_id'], 'test': True, 'check': ['customer_code', 'emp_id',]},
                         'fAMCInv': {'cid': ['1'], 'usecols': ['customer_code', 'order_id', 'emp_id'], 'test': True, 'check': ['customer_code', 'order_id', 'emp_id']},
                         'fCC': {'cid': ['1', '2',], 'usecols': ['emp_id'], 'test': True, 'check': ['emp_id']},
                         'fEOS': {'cid': ['1'], 'usecols': ['emp_id'], 'test': True, 'check': ['emp_id']},
                         'fER': {'cid': ['1'], 'usecols': ['emp_id'], 'test': True, 'check': ['emp_id']},
                         'fLeave': {'cid': ['1'], 'usecols': ['emp_id'], 'test': True, 'check': ['emp_id']},
                         'fCreditNote': {'cid': ['1', '4', '6', '5', '8', '2',], 'usecols': ['ledger_code', 'order_id'], 'test': True, 'check': ['ledger_code', 'order_id']},
                         'fMI': {'cid': ['1', '2',], 'usecols': ['part_number', 'emp_id','order_id'], 'test': True, 'check': ['part_number', 'emp_id']},
                         'fOT': {'cid': ['1', '2',], 'usecols': ['emp_id'], 'test': True, 'check': ['emp_id']},
                         'fOutSourceInv': {'cid': ['1', '2',], 'usecols': ['order_id', 'customer_code'], 'test': True, 'check': ['order_id', 'customer_code']},
                         'fProInv': {'cid': ['1', '6', '5', '8',], 'usecols': ['customer_code', 'order_id', 'emp_id'], 'test': True, 'check': ['customer_code', 'order_id', 'emp_id']},
                         'fTkt': {'cid': ['1'], 'usecols': ['emp_id'], 'test': True, 'check': ['emp_id']},
                         'fRentInv': {'cid': ['4'], 'usecols': ['order_id', 'customer_code'], 'test': False},
                         'fPurchase': {'cid': ['4', '6', '5', '8', '2',], 'usecols': ['ledger_code'], 'test': False},
                         'fInvCI': {'cid': ['2',], 'usecols': ['customer_code', 'emp_id'], 'test': False}, }

# this will returns set containing sheet names to check for the selected company
sheets: set[str] = {i for i in sheets_to_check if str(
    company_id+1) in sheets_to_check[i]['cid']}

# returns a dictionary containing dataframe as key and columns to be tested in that dataframe as value for the selected company
checks: list[str] = {sheet: sheets_to_check[sheet]['check']
                     for sheet in sheets if sheets_to_check[sheet]['test']}
dataframes = {}
for sheet in sheets:
    print(f'Now reading {sheet}...')
    # read the sheets from the workbook for the selected company and store them in a dictionary
    dataframes[sheet] = pd.read_excel(
        io=file_path, sheet_name=sheet, usecols=sheets_to_check[sheet]['usecols'])

# the fMI dataframe read from the work book requirs some cleaning and filling of missing values
fMI: pd.DataFrame = dataframes.get('fMI', pd.DataFrame(data={'part_number':['-'],'order_id':['-'],'emp_id':['-']}))
fMI.fillna(value='-', inplace=True)
fMI['order_id'] = fMI['order_id'].str.split('-', expand=True)[0]
fMI['order_id'] = fMI['order_id'].fillna('PH/CTR230020')
fMI['emp_id'] = fMI['emp_id'].str.split('-', expand=True)[0]
fMI['emp_id'] = fMI['emp_id'].fillna('PH00001')
# reassing the cleaned dataframe to the dictionary
dataframes['fMI'] = fMI

# orders in the Adler system are stored as multiple tables. If the company being read has any of those worksheets, such data being read
# must be cleaned prior to be used. 

# for NBN Logistics 
dLogContract = [job.split(sep='-')[0] for job in dataframes.get('dLogContract', pd.DataFrame(data={'order_id': ['-']}))['order_id'].unique()]

# for ESS
dOrderAMC = [job.split(sep='-')[0] for job in dataframes.get('dOrderAMC', pd.DataFrame(data={'order_id': ['-']}))['order_id'].unique()]

# for NBNH/ESS/NBN Tech/ NBN Triptikum and GAT
dCusOrder = [job.split(sep='-')[0] for job in dataframes.get('dCusOrder', pd.DataFrame(data={'order_id': ['-']}))['order_id'].unique()]

# for ESS/ PH
dContract = [job.split(sep='-')[0] for job in dataframes.get('dContract', pd.DataFrame(data={'order_id': ['-']}))['order_id'].unique()]

# for NBN Real estate
dRentContract = [job.split(sep='-')[0] for job in dataframes.get('dRentContract', pd.DataFrame(data={'order_id': ['-']}))['order_id'].unique()]
dJobs = dOrderAMC + dCusOrder + dContract + dLogContract + dRentContract

# the base_lists dictionary contains the unique values of the primary keys from the master list
base_lists: dict = {'ledger_code': dataframes.get('dCoAAdler')['ledger_code'].unique(),
                    'order_id': dJobs,
                    'emp_id': dataframes.get('dEmployee')['emp_id'].unique(),
                    'room_id': dataframes.get('dRoom',pd.DataFrame(data={'room_id':[np.nan]}))['room_id'].unique(),
                    'part_number': dataframes.get('dStock',pd.DataFrame(data={'part_number':[np.nan]}))['part_number'].unique(),
                    'customer_code': dataframes.get('dCustomer', pd.DataFrame(data={'customer_code': [np.nan]}))['customer_code'].unique(),
                    'service_element_code': dataframes.get('dServiceTypes', pd.DataFrame(data={'service_element_code': [np.nan]}))['service_element_code'].unique(), }


def missing_data(dataframe: pd.DataFrame, tests: list) -> list:
    """for a given dataframe and columns to tested in that dataframe checks whether foreign keys under columns to be tested are present in the base list. 

    Args:
        dataframe (pd.DataFrame): a dataframe to be tested
        tests (list): columns to be tested in the dataframe

    Returns:
        list: list of foreign keys which are not present in the base list
    """
    for test in tests:
        compare_list: list = list(
            set(dataframes.get(dataframe)[f'{test}'].tolist()))
        missing_values: list = [
            item for item in compare_list if item not in base_lists[test]]
        if not missing_values:
            print(f'{dataframe}:{test}-{Fore.GREEN}PASSED{Fore.RESET}')
        else:
            print(
                f'{dataframe}:{test}-{Fore.RED}FAILED{Fore.RESET}\nPlease check {missing_values}')


print('\n')  # to leave a space after Now reading fGL...

for dataframe, tests in checks.items():
    missing_data(dataframe=dataframe, tests=tests)
