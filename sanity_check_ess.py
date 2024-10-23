import pandas as pd
from colorama import Fore, init

init()

file_path = r'C:\Masters\Data-ESS.xlsx'

dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['ledger_code'],dtype={'ledger_code': 'str'})
dCusOrder: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCusOrder', usecols=['order_id', 'customer_code','emp_id'])
dContract: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dContract',
                                         usecols=['order_id', 'customer_code', 'emp_id'])
dContract['order_id'] = dContract['order_id'].str.split('-', expand=True)[0].str.strip()
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['emp_id'])
dCustomer: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomer',
                                         usecols=['customer_code', 'ledger_code'])
dStock: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dStock', usecols=['part_number'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['ledger_code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['ledger_code'],dtype={'ledger_code': 'str'})
fOutSourceInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOutSourceInv',
                                            usecols=['order_id', 'customer_code'])
fAMCInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fAMCInv', usecols=['customer_code','order_id','emp_id'])
fProInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fProInv', usecols=['customer_code', 'order_id','emp_id'])
fCreditNote: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCreditNote', usecols=['ledger_code','order_id'])
fMI: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fMI', usecols=['part_number', 'emp_id'])
fMI['emp_id'] = fMI['emp_id'].str.split('-', expand=True)[0]
fBudget: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fBudget', usecols=['ledger_code'])
fOT: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOT', usecols=['emp_id'])
fCC: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCC', usecols=['emp_id'])
fLeave: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fLeave', usecols=['emp_id'])
fTkt: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fTkt', usecols=['emp_id'])
fER: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fER', usecols=['emp_id'])
fEOS: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fEOS', usecols=['emp_id'])
dOrderAMC: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dOrderAMC', usecols=['customer_code','emp_id','order_id'])
dJobs:pd.DataFrame = pd.concat([dOrderAMC,dCusOrder,dContract])


dataframes: dict = {'dCoAAdler': dCoAAdler, 'fOT': fOT, 'dContracts': dContract,'fTkt':fTkt,
                    'dEmployee': dEmployee, 'dCustomers': dCustomer,'fLeave':fLeave,'fER':fER,'fEOS':fEOS,
                    'fGL': fGL, 'fCollection': fCollection, 'dStock': dStock, 'fCreditNote': fCreditNote,
                    'fCC': fCC, 'fBudget': fBudget, 'fMI': fMI, 'dCusOrder': dCusOrder, 'fAMCInv': fAMCInv,
                    'fProInv': fProInv,'fOutSourceInv': fOutSourceInv,'dOrderAMC':dOrderAMC}

checks: dict = {
    'fCC': ['emp_id'],
    'fGL': ['ledger_code'],
    'dContracts': ['customer_code', 'emp_id'],
    'fCollection': ['ledger_code'],
    'fCreditNote': ['ledger_code','order_id'],
    'fMI': ['emp_id', 'part_number'],
    'fOT': ['emp_id'],
    'fOutSourceInv': ['customer_code', 'order_id'],
    'fAMCInv': ['customer_code','order_id','emp_id'],
    'fProInv': ['customer_code', 'order_id','emp_id'],
    'fBudget': ['ledger_code'],
    'fLeave':['emp_id'],
    'fTkt':['emp_id'],
    'fER':['emp_id'],
    'fEOS':['emp_id'],
    'dCusOrder':['customer_code','emp_id'],
    'dOrderAMC':['customer_code','emp_id']
}

base_lists: dict = {'ledger_code': set(dCoAAdler['ledger_code'].tolist()),
                    'emp_id': set(dEmployee['emp_id'].tolist()),
                    'customer_code': set(dCustomer['customer_code'].tolist()),
                    'part_number': set(dStock['part_number'].tolist()),
                    'order_id': set(dJobs['order_id'].tolist()),
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
