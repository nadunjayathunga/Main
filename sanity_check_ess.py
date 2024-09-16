import pandas as pd
from colorama import Fore, init

init()

file_path = r'C:\Masters\Data-ESS.xlsx'

dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['Ledger_Code'],dtype={'Ledger_Code': 'str'})
dCusOrder: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCusOrder', usecols=['Order_ID', 'Customer_Code','Employee_Code'])
dContract: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dContract',
                                         usecols=['Order_ID', 'Customer_Code', 'Employee_Code'])
dContract['Order_ID'] = dContract['Order_ID'].str.split('-', expand=True)[0].str.strip()
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['Employee_Code'])
dCustomer: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomer',
                                         usecols=['Customer_Code', 'Ledger_Code'])
dStock: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dStock', usecols=['Part Number'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['Ledger_Code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['Ledger_Code'],dtype={'Ledger_Code': 'str'})
fOutSourceInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOutSourceInv',
                                            usecols=['Order_ID', 'Customer_Code'])
fAMCInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fAMCInv', usecols=['Customer_Code','Order_ID','Employee_Code'])
fProInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fProInv', usecols=['Customer_Code', 'Order_ID','Employee_Code'])
fCreditNote: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCreditNote', usecols=['Ledger_Code','Order_ID'])
fMI: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fMI', usecols=['Part Number', 'Employee_Code'])
fMI['Employee_Code'] = fMI['Employee_Code'].str.split('-', expand=True)[0]
fBudget: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fBudget', usecols=['Ledger_Code'])
fOT: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOT', usecols=['Employee_Code'])
fCC: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCC', usecols=['Employee_Code'])
fLeave: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fLeave', usecols=['Employee_Code'])
fTkt: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fTkt', usecols=['Employee_Code'])
fER: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fER', usecols=['Employee_Code'])
fEOS: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fEOS', usecols=['Employee_Code'])
dOrderAMC: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dOrderAMC', usecols=['Customer_Code','Employee_Code','Order_ID'])
dJobs:pd.DataFrame = pd.concat([dOrderAMC,dCusOrder,dContract])


dataframes: dict = {'dCoAAdler': dCoAAdler, 'fOT': fOT, 'dContracts': dContract,'fTkt':fTkt,
                    'dEmployee': dEmployee, 'dCustomers': dCustomer,'fLeave':fLeave,'fER':fER,'fEOS':fEOS,
                    'fGL': fGL, 'fCollection': fCollection, 'dStock': dStock, 'fCreditNote': fCreditNote,
                    'fCC': fCC, 'fBudget': fBudget, 'fMI': fMI, 'dCusOrder': dCusOrder, 'fAMCInv': fAMCInv,
                    'fProInv': fProInv,'fOutSourceInv': fOutSourceInv,'dOrderAMC':dOrderAMC}

checks: dict = {
    'fCC': ['Employee_Code'],
    'fGL': ['Ledger_Code'],
    'dContracts': ['Customer_Code', 'Employee_Code'],
    'fCollection': ['Ledger_Code'],
    'fCreditNote': ['Ledger_Code','Order_ID'],
    'fMI': ['Employee_Code', 'Part Number'],
    'fOT': ['Employee_Code'],
    'fOutSourceInv': ['Customer_Code', 'Order_ID'],
    'fAMCInv': ['Customer_Code','Order_ID','Employee_Code'],
    'fProInv': ['Customer_Code', 'Order_ID','Employee_Code'],
    'fBudget': ['Ledger_Code'],
    'fLeave':['Employee_Code'],
    'fTkt':['Employee_Code'],
    'fER':['Employee_Code'],
    'fEOS':['Employee_Code'],
    'dCusOrder':['Customer_Code','Employee_Code'],
    'dOrderAMC':['Customer_Code','Employee_Code']
}

base_lists: dict = {'Ledger_Code': set(dCoAAdler['Ledger_Code'].tolist()),
                    'Employee_Code': set(dEmployee['Employee_Code'].tolist()),
                    'Customer_Code': set(dCustomer['Customer_Code'].tolist()),
                    'Part Number': set(dStock['Part Number'].tolist()),
                    'Order_ID': set(dJobs['Order_ID'].tolist()),
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
