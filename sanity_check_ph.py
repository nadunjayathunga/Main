import pandas as pd
from colorama import Fore, init
from data import company_info
import math

init()

print('\n')
for idx, company in enumerate(company_info):
    print(f'{idx}\t{company["data"]["long_name"]}')
print('\n')
company_id: int = int(input('Please enter company ID>> '))
print('\n')

file_name: str = company_info[company_id]['data']['file_name']
file_path = f'C:\Masters\{file_name}.xlsx'

dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['ledger_code'])
dJobs: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dContract',
                                    usecols=['order_id', 'customer_code', 'emp_id'])
dJobs['order_id'] = dJobs['order_id'].str.split('-', expand=True)[0]
dCustomer: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomer',
                                         usecols=['customer_code', 'ledger_code'])
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['emp_id'])
dStock: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dStock', usecols=['part_number'])
fOT: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOT', usecols=['emp_id'])
fCC: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCC', usecols=['emp_id'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['ledger_code'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['ledger_code'])
fTimesheet: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fTimesheet', usecols=['employee_code'])
fCreditNote: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCreditNote', usecols=['ledger_code', 'order_id'])
fOutSourceInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOutSourceInv', usecols=['order_id', 'customer_code'])
fInvCI: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fInvCI',
                                     usecols=['customer_code', 'emp_id'])
fMI: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fMI', usecols=['part_number', 'emp_id', 'order_id'])
fMI['order_id'] = fMI['order_id'].str.split('-', expand=True)[0]
fMI['order_id'] = fMI['order_id'].fillna('PH/CTR230020')
fMI['emp_id'] = fMI['emp_id'].str.split('-', expand=True)[0]
fMI['emp_id'] = fMI['emp_id'].fillna('PH00001')

dataframes: dict = {'dCoAAdler': dCoAAdler, 'fOT': fOT, 'dJobs': dJobs,
                    'dEmployee': dEmployee, 'dCustomers': dCustomer,
                    'fGL': fGL, 'fCollection': fCollection, 'dStock': dStock, 'fCC': fCC, 'fOutSourceInv': fOutSourceInv,
                    'fInvCI': fInvCI, 'fTimesheet': fTimesheet, 'fCreditNote': fCreditNote, 'fMI': fMI}

column_rename: dict = {'ledger_code': ['ledger_code', 'ledger_code'],
                       'order_id': ['order_id', 'order_id', 'Order Reference Number',
                                                  'order_id'],
                       'customer_code': ['customer_code', 'customer_code'],
                       'part_number': ['part_number'],
                       'emp_id': ['emp_id', 'emp_id', 'cost_center', 'emp_id', 'employee_code',
                                         'emp_id', 'emp_id'],
                       }

checks: dict = {'fTimesheet': ['emp_id'],
                'fCC': ['emp_id'],
                'fGL': ['ledger_code'],
                'dJobs': ['customer_code', 'emp_id'],
                'fCollection': ['ledger_code'],
                'fCreditNote': ['ledger_code', 'order_id'],
                'fOutSourceInv': ['order_id', 'customer_code'],
                'fInvCI': ['customer_code', 'emp_id'],
                'fMI': ['order_id', 'emp_id', 'part_number'],
                'fOT': ['emp_id']
                }

base_lists: dict = {'ledger_code': list(set(dCoAAdler['ledger_code'].tolist())),
                    'order_id': [job.split(sep='-')[0] for job in
                                               list(set(dJobs['order_id'].tolist()))],
                    'emp_id': (set(dEmployee['emp_id'].tolist())),
                    'customer_code': list(set(dCustomer['customer_code'].tolist())),
                    'part_number': list(set(dStock['part_number'].tolist())),
                    }


def missing_data(dataframe: pd.DataFrame, tests: list) -> list:
    for test in tests:
        column_name = [column_name for column_name in list(dataframes.get(dataframe).columns) if
                       column_name in column_rename.get(test)][0]
        compare_list: list = list(set(dataframes.get(dataframe)[f'{column_name}'].tolist()))
        missing_values: list = [item for item in compare_list if item not in base_lists[test]]
        if not missing_values:
            print(f'{dataframe}:{test}-{Fore.GREEN}PASSED{Fore.RESET}')
        else:
            print(f'{dataframe}:{test}-{Fore.RED}FAILED{Fore.RESET}\nPlease check {missing_values}')


for dataframe, tests in checks.items():
    missing_data(dataframe=dataframe, tests=tests)
