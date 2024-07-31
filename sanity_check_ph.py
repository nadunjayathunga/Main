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

dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['Ledger_Code'])
dJobs: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dJobs',
                                    usecols=['Order_Reference_Number', 'Customer_Code', 'Emp_id'])
dJobs['Order_Reference_Number'] = dJobs['Order_Reference_Number'].str.split('-', expand=True)[0]
dCustomers: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomers',
                                         usecols=['Customer_Code', 'Ledger_Code'])
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['Employee_Code'])
dStock: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dStock', usecols=['Part Number'])
fOT: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOT', usecols=['cost_center'])
fCC: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCC', usecols=['Emp_Code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['Ledger Code', ])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['Ledger Code'])
fTimesheet: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fTimesheet', usecols=['employee_code'])
fCN: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCN', usecols=['Ledger Code', 'Job Code'])
fInvoices: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fInvoices', usecols=['Job Code', 'Customer_Code'])
fInvCI: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fInvCI',
                                     usecols=['Customer Code', 'Sales Engineer Code'])
fMI: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fMI', usecols=['Part Number', 'Cost Centre', 'Job'])
fMI['Job'] = fMI['Job'].str.split('-', expand=True)[0]
fMI['Job'] = fMI['Job'].fillna('PH/CTR230020')
fMI['Cost Centre'] = fMI['Cost Centre'].str.split('-', expand=True)[0]
fMI['Cost Centre'] = fMI['Cost Centre'].fillna('PH00001')

dataframes: dict = {'dCoAAdler': dCoAAdler, 'fOT': fOT, 'dJobs': dJobs,
                    'dEmployee': dEmployee, 'dCustomers': dCustomers,
                    'fGL': fGL, 'fCollection': fCollection, 'dStock': dStock, 'fCC': fCC, 'fInvoices': fInvoices,
                    'fInvCI': fInvCI, 'fTimesheet': fTimesheet, 'fCN': fCN, 'fMI': fMI}

column_rename: dict = {'Ledger_Code': ['Ledger_Code', 'Ledger Code'],
                       'Order_Reference_Number': ['Order_Reference_Number', 'Job Code', 'Order Reference Number',
                                                  'Job'],
                       'Customer_Code': ['Customer_Code', 'Customer Code'],
                       'Part Number': ['Part Number'],
                       'Employee_Code': ['Employee_Code', 'Emp_id', 'cost_center', 'Emp_Code', 'employee_code',
                                         'Sales Engineer Code', 'Cost Centre'],
                       }

checks: dict = {'fTimesheet': ['Employee_Code'],
                'fCC': ['Employee_Code'],
                'fGL': ['Ledger_Code'],
                'dJobs': ['Customer_Code', 'Employee_Code'],
                'fCollection': ['Ledger_Code'],
                'fCN': ['Ledger_Code', 'Order_Reference_Number'],
                'fInvoices': ['Order_Reference_Number', 'Customer_Code'],
                'fInvCI': ['Customer_Code', 'Employee_Code'],
                'fMI': ['Order_Reference_Number', 'Employee_Code', 'Part Number'],
                'fOT': ['Employee_Code']
                }

base_lists: dict = {'Ledger_Code': list(set(dCoAAdler['Ledger_Code'].tolist())),
                    'Order_Reference_Number': [job.split(sep='-')[0] for job in
                                               list(set(dJobs['Order_Reference_Number'].tolist()))],
                    'Employee_Code': (set(dEmployee['Employee_Code'].tolist())),
                    'Customer_Code': list(set(dCustomers['Customer_Code'].tolist())),
                    'Part Number': list(set(dStock['Part Number'].tolist())),
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
