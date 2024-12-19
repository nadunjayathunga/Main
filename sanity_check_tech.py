import pandas as pd
from colorama import Fore, init

init()

file_path = r'C:\Masters\Data-NBNT.xlsx'

dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['ledger_code'],dtype={'ledger_code': 'str'})
dContract: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dContract',usecols=['order_id', 'customer_code','emp_id'])
dCustomer: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomer',usecols=['customer_code', 'ledger_code'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['ledger_code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['ledger_code'],dtype={'ledger_code': 'str'})
fInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fInv',usecols=['order_id', 'customer_code'])
fCreditNote: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCreditNote', usecols=['ledger_code'])
fAP: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fAP', usecols=['ledger_code'])
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['emp_id'])

dataframes: dict = {'dCoAAdler': dCoAAdler, 'dContract': dContract,'dCustomers': dCustomer,
                    'fGL': fGL, 'fCollection': fCollection,  'fCreditNote': fCreditNote,'fInv':fInv,'fAP':fAP ,'dEmployee':dEmployee}

checks: dict = {
    'fGL': ['ledger_code'],
    'dContract': ['customer_code','emp_id','order_id'],
    'fCollection': ['ledger_code'],
    'fCreditNote': ['ledger_code'],
    'fInv': ['customer_code', 'order_id'],
    'fAP':['ledger_code'],
    # 'dEmployee':['emp_id']
}

base_lists: dict = {'ledger_code': set(dCoAAdler['ledger_code'].tolist()),
                    'customer_code': set(dCustomer['customer_code'].tolist()),
                    'order_id': set(dContract['order_id'].tolist()),
                    'emp_id': set(dEmployee['emp_id'].tolist()),
                    }

def missing_data(dataframe: pd.DataFrame, tests: list) -> list:
    for test in tests:
        compare_list: list = [i for i in set(dataframes.get(dataframe)[f'{test}'].tolist()) if
                              isinstance(i, str)]
        missing_values: list = [item for item in compare_list if item not in base_lists[test]]
        if not missing_values:
            print(f'{dataframe}:{test}-{Fore.GREEN}PASSED{Fore.RESET}')
        else:
            print(f'{dataframe}:{test}-{Fore.RED}FAILED{Fore.RESET}\nPlease check {missing_values}')


for dataframe, tests in checks.items():
    missing_data(dataframe=dataframe, tests=tests)
