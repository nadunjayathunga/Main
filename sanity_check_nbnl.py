import pandas as pd
from colorama import Fore, init
from data import company_info

init()

print('\n')
for idx, company in enumerate(company_info):
    print(f'{idx}\t{company["data"]["long_name"]}')
print('\n')
company_id: int = int(input('Please enter company ID>> '))
print('\n')

file_name: str = company_info[company_id]['data']['file_name']
file_path = f'C:\Masters\{file_name}.xlsx'

fBudget: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fBudget', usecols=['L5-Code'])
dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['Ledger_Code'])
fData: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fData',
                                    usecols=['Ledger_Code', 'Job_Code', 'Service Element Code'])
dJobs: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dJobs', usecols=['Job_Number', 'Customer_Code', 'emp_id'])
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['emp_id'])
dCustomers: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomers', usecols=['Customer_Code'])
dServiceTypes: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dServiceTypes', usecols=['Service Code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['Ledger Code', ])
fNotInvoiced: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fNotInvoiced',
                                           usecols=['Job Number', 'Ledger Code'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['Ledger Code'])
fLogInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fLogInv',
                                      usecols=['Job Number', 'Customer Code', 'Sales Person Code'])

dataframes: dict = {'fBudget': fBudget, 'dCoAAdler': dCoAAdler, 'fData': fData, 'dJobs': dJobs,
                    'dEmployee': dEmployee, 'dCustomers': dCustomers, 'dServiceTypes': dServiceTypes,
                    'fGL': fGL, 'fNotInvoiced': fNotInvoiced, 'fCollection': fCollection, 'fLogInv': fLogInv}

column_rename: dict = {'Ledger_Code': ['L5-Code', 'Ledger_Code', 'Ledger Code', 'GL'],
                       'Job_Number': ['Job_Number', 'Job_Code', 'Job Number'],
                       'Service Element Code': ['Service Element Code', 'Service Code'],
                       'Customer_Code': ['Customer_Code', 'Customer Code'],
                       'emp_id': ['emp_id', 'Sales Person Code']
                       }

checks: dict = {'fBudget': ['Ledger_Code'],
                'fData': ['Ledger_Code', 'Job_Number', 'Service Element Code'],
                'fGL': ['Ledger_Code'],
                'dJobs': ['Customer_Code', 'emp_id'],
                'fNotInvoiced': ['Job_Number', 'Ledger_Code'],
                'fCollection': ['Ledger_Code'],
                'fLogInv': ['Job_Number', 'Customer_Code', 'emp_id']
                }

base_lists: dict = {'Ledger_Code': list(set(dCoAAdler['Ledger_Code'].tolist())),
                    'Job_Number': [job.split(sep='-')[0] for job in list(set(dJobs['Job_Number'].tolist()))],
                    'emp_id': list(set(dEmployee['emp_id'].tolist())),
                    'Customer_Code': list(set(dCustomers['Customer_Code'].tolist())),
                    'Service Element Code': list(set(dServiceTypes['Service Code'].tolist())), }


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
