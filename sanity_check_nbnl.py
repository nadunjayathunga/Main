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

fBudget: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fBudget', usecols=['ledger_code'])
dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['ledger_code'])
fData: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fData',
                                    usecols=['ledger_code', 'order_id', 'service_element_code'])
dLogContract: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dLogContract', usecols=['order_id', 'customer_code', 'emp_id'])
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['emp_id'])
dCustomer: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomer', usecols=['customer_code'])
dServiceTypes: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dServiceTypes', usecols=['service_element_code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['ledger_code', ])
fNotInvoiced: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fNotInvoiced',
                                           usecols=['order_id', 'ledger_code'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['ledger_code'])
fLogInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fLogInv',
                                      usecols=['order_id', 'customer_code', 'emp_id'])

dataframes: dict = {'fBudget': fBudget, 'dCoAAdler': dCoAAdler, 'fData': fData, 'dLogContract': dLogContract,
                    'dEmployee': dEmployee, 'dCustomer': dCustomer, 'dServiceTypes': dServiceTypes,
                    'fGL': fGL, 'fNotInvoiced': fNotInvoiced, 'fCollection': fCollection, 'fLogInv': fLogInv}

checks: dict = {'fBudget': ['ledger_code'],
                'fData': ['ledger_code', 'order_id', 'service_element_code'],
                'fGL': ['ledger_code'],
                'dLogContract': ['customer_code', 'emp_id'],
                'fNotInvoiced': ['order_id', 'ledger_code'],
                'fCollection': ['ledger_code'],
                'fLogInv': ['order_id', 'customer_code', 'emp_id']
                }

base_lists: dict = {'ledger_code': dCoAAdler['ledger_code'].unique(),
                    'order_id': [job.split(sep='-')[0] for job in dLogContract['order_id'].unique()],
                    'emp_id': dEmployee['emp_id'].unique(),
                    'customer_code': dCustomer['customer_code'].unique(),
                    'service_element_code': dServiceTypes['service_element_code'].unique(), }


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
