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

fBudget: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fBudget', usecols=['Ledger_Code'])
dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['Ledger_Code'])
fData: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fData',
                                    usecols=['Ledger_Code', 'Job_Number', 'Service_Code'])
dLogContract: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dLogContract', usecols=['Job_Number', 'Customer_Code', 'Employee_Code'])
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['Employee_Code'])
dCustomer: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomer', usecols=['Customer_Code'])
dServiceTypes: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dServiceTypes', usecols=['Service_Code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['Ledger_Code', ])
fNotInvoiced: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fNotInvoiced',
                                           usecols=['Job_Number', 'Ledger_Code'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['Ledger_Code'])
fLogInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fLogInv',
                                      usecols=['Job_Number', 'Customer_Code', 'Employee_Code'])

dataframes: dict = {'fBudget': fBudget, 'dCoAAdler': dCoAAdler, 'fData': fData, 'dLogContract': dLogContract,
                    'dEmployee': dEmployee, 'dCustomer': dCustomer, 'dServiceTypes': dServiceTypes,
                    'fGL': fGL, 'fNotInvoiced': fNotInvoiced, 'fCollection': fCollection, 'fLogInv': fLogInv}

checks: dict = {'fBudget': ['Ledger_Code'],
                'fData': ['Ledger_Code', 'Job_Number', 'Service_Code'],
                'fGL': ['Ledger_Code'],
                'dLogContract': ['Customer_Code', 'Employee_Code'],
                'fNotInvoiced': ['Job_Number', 'Ledger_Code'],
                'fCollection': ['Ledger_Code'],
                'fLogInv': ['Job_Number', 'Customer_Code', 'Employee_Code']
                }

base_lists: dict = {'Ledger_Code': dCoAAdler['Ledger_Code'].unique(),
                    'Job_Number': [job.split(sep='-')[0] for job in dLogContract['Job_Number'].unique()],
                    'Employee_Code': dEmployee['Employee_Code'].unique(),
                    'Customer_Code': dCustomer['Customer_Code'].unique(),
                    'Service_Code': dServiceTypes['Service_Code'].unique(), }


def missing_data(dataframe: pd.DataFrame, tests: list) -> list:
    for test in tests:
        compare_list: list = list(set(dataframes.get(dataframe)[f'{test}'].tolist()))
        missing_values: list = [item for item in compare_list if item not in base_lists[test]]
        if not missing_values:
            print(f'{dataframe}:{test}-{Fore.GREEN}PASSED{Fore.RESET}')
        else:
            print(f'{dataframe}:{test}-{Fore.RED}FAILED{Fore.RESET}\nPlease check {missing_values}')


for dataframe, tests in checks.items():
    missing_data(dataframe=dataframe, tests=tests)
