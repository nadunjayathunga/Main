import pandas as pd
from colorama import Fore, init

init()

file_path = r'C:\Masters\Data-NBNRE.xlsx'

dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['ledger_code'],dtype={'ledger_code': 'str'})
dRentContract: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dRentContract',usecols=['order_id', 'customer_code','room_id'])
dCustomer: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomer',usecols=['customer_code', 'ledger_code'])
dStock: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dStock', usecols=['part_number'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['ledger_code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['ledger_code'],dtype={'ledger_code': 'str'})
fInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fInv',usecols=['order_id', 'customer_code'])
fCreditNote: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCreditNote', usecols=['ledger_code'])
dRoom: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dRoom', usecols=['room_id'])
fAP: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fAP', usecols=['ledger_code'])

dataframes: dict = {'dCoAAdler': dCoAAdler, 'dRentContract': dRentContract,'dCustomers': dCustomer,
                    'fGL': fGL, 'fCollection': fCollection, 'dStock': dStock, 'fCreditNote': fCreditNote,'fInv':fInv,'dRoom':dRoom,'fAP':fAP }

checks: dict = {
    'fGL': ['ledger_code'],
    'dRentContract': ['customer_code','room_id','order_id'],
    'fCollection': ['ledger_code'],
    'fCreditNote': ['ledger_code'],
    'fInv': ['customer_code', 'order_id'],
    'fAP':['ledger_code']
}

base_lists: dict = {'ledger_code': set(dCoAAdler['ledger_code'].tolist()),
                    'customer_code': set(dCustomer['customer_code'].tolist()),
                    'part_number': set(dStock['part_number'].tolist()),
                    'order_id': set(dRentContract['order_id'].tolist()),
                    'room_id': set(dRoom['room_id'].tolist()),
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
