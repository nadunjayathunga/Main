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

dCoAAdler: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCoAAdler', usecols=['Ledger_Code'],dtype={'Ledger_Code': 'str'})
dCusOrder: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCusOrder', usecols=['Order_ID', 'Customer Code'])
dContracts: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dContracts',
                                         usecols=['Order_Reference_Number', 'Customer_Code', 'Emp_id'])
dEmployee: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dEmployee', usecols=['Employee_Code'])
dCustomers: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dCustomers',
                                         usecols=['Customer_Code', 'Ledger_Code'])
dStock: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='dStock', usecols=['Part Number'])
fCollection: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCollection', usecols=['Ledger Code'])
fGL: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fGL', usecols=['Ledger Code'],dtype={'Ledger Code': 'str'})
fOutSourceInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOutSourceInv',
                                            usecols=['Job_id', 'Customer_Code'])
fAMCInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fAMCInv', usecols=['Customer_Code'])
fProInv: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fProInv', usecols=['Customer_Code', 'Order_ID','Sales Engineer Code'])
fCreditNote: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCreditNote', usecols=['Ledger_Code'])
fMI: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fMI', usecols=['Part Number', 'Cost Centre'])
fMI['Cost Centre'] = fMI['Cost Centre'].str.split('-', expand=True)[0]
fBudget: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fBudget', usecols=['L5-Code'])
fOT: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fOT', usecols=['cost_center'])
fCC: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fCC', usecols=['Emp_Code'])
fLeave: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fLeave', usecols=['Cost Center Code  :'])
fTkt: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fTkt', usecols=['Cost Center Code  :'])
fER: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fER', usecols=['Cost Center Code  :'])
fEOS: pd.DataFrame = pd.read_excel(io=file_path, sheet_name='fEOS', usecols=['Cost Center Code  :'])


dataframes: dict = {'dCoAAdler': dCoAAdler, 'fOT': fOT, 'dContracts': dContracts,'fTkt':fTkt,
                    'dEmployee': dEmployee, 'dCustomers': dCustomers,'fLeave':fLeave,'fER':fER,'fEOS':fEOS,
                    'fGL': fGL, 'fCollection': fCollection, 'dStock': dStock, 'fCreditNote': fCreditNote,
                    'fCC': fCC, 'fBudget': fBudget, 'fMI': fMI, 'dCusOrder': dCusOrder, 'fAMCInv': fAMCInv,
                    'fProInv': fProInv,'fOutSourceInv': fOutSourceInv,}

column_rename: dict = {'Ledger_Code': ['Ledger_Code', 'Ledger Code', 'L5-Code'],
                       'Order_Reference_Number': ['Order_Reference_Number', 'Job_id', ],
                       'Customer_Code': ['Customer_Code', 'Customer Code'],
                       'Part Number': ['Part Number'],
                       'Employee_Code': ['Employee_Code', 'Emp_id', 'Emp_Code', 'cost_center', 'Cost Centre','Sales Engineer Code','Cost Center Code  :'],
                       'Order_ID': ['Order_ID']
                       }

checks: dict = {
    'fCC': ['Employee_Code'],
    'fGL': ['Ledger_Code'],
    'dContracts': ['Customer_Code', 'Employee_Code'],
    'fCollection': ['Ledger_Code'],
    'fCreditNote': ['Ledger_Code'],
    'fMI': ['Employee_Code', 'Part Number'],
    'fOT': ['Employee_Code'],
    'fOutSourceInv': ['Customer_Code', 'Order_Reference_Number'],
    'fAMCInv': ['Customer_Code'],
    'fProInv': ['Customer_Code', 'Order_ID','Employee_Code'],
    'fBudget': ['Ledger_Code'],
    'fLeave':['Employee_Code'],
    'fTkt':['Employee_Code'],
    'fER':['Employee_Code'],
    'fEOS':['Employee_Code'],
}

base_lists: dict = {'Ledger_Code': set(dCoAAdler['Ledger_Code'].tolist()),
                    'Order_Reference_Number': set(dContracts['Order_Reference_Number'].tolist()),
                    'Employee_Code': set(dEmployee['Employee_Code'].tolist()),
                    'Customer_Code': set(dCustomers['Customer_Code'].tolist()),
                    'Part Number': set(dStock['Part Number'].tolist()),
                    'Order_ID': set(dCusOrder['Order_ID'].tolist()),
                    }

def missing_data(dataframe: pd.DataFrame, tests: list) -> list:
    for test in tests:
        column_name = [column_name for column_name in list(dataframes.get(dataframe).columns) if
                       column_name in column_rename.get(test)][0]
        compare_list: list = [i for i in set(dataframes.get(dataframe)[f'{column_name}'].tolist()) if
                              isinstance(i, str)]
        missing_values: list = [item for item in compare_list if item not in base_lists[test]]
        if not missing_values:
            print(f'{dataframe}:{test}-{Fore.GREEN}PASSED{Fore.RESET}')
        else:
            print(f'{dataframe}:{test}-{Fore.RED}FAILED{Fore.RESET}\nPlease check {missing_values}')


for dataframe, tests in checks.items():
    missing_data(dataframe=dataframe, tests=tests)
