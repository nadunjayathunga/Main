import pandas as pd
from datetime import datetime,timedelta
import numpy as np
from dateutil.relativedelta import relativedelta
import matplotlib.pyplot as plt

PATH = r'C:\Users\NadunJayathunga\OneDrive - NBN Holdings\Financials\Other\Programmes\Dashboards\Main\test.csv'
# start_date = datetime(year=2023,month=10,day=1)
end_date = datetime(year=2024,month=9,day=30)
# df = pd.read_csv(PATH)
# df['termination_date'] = pd.to_datetime(df['termination_date'])
# df = df.loc[((df['termination_date']>=end_date) | df['termination_date'].isna()) & df['designation'].str.contains(pat='Sales',case=False) ,['emp_id','designation','termination_date']]

# print(df)
special_emp:dict = {'NBNL0088':{'ticket':{'self':0,
                                                'spouse':1,
                                                'dependent':2},
                                    'insurance':{'cat':'b',
                                                 'qty':{'self':1,
                                                        'spouse':1,
                                                        'dependent':2}},
                                    'rp':{'self':1,
                                                'spouse':1,
                                                'dependent':2}
                                    }
                    }

ctc_amount = {'insurance':
              {'a1': {'adult': 9_146, 'dependent': 5_663},
               'a2': {'adult': 6_243, 'dependent': 0},
               'b': {'adult': 5_898, 'dependent': 3_693},
               'c': {'adult': 2_788, 'dependent': 2_182},
               'd': {'adult': 1_977, 'dependent': 0},
               },
              'rp':{'adult': 1_220, 'minor': 500}}

def other_benefits(emp_id:str)->float:
    emp_dict:dict =  special_emp[emp_id]
    ticket_total:float = sum([v for v  in emp_dict['ticket'].values()]) * ticket_allowance

emp_id = 'NBNL0088'
ticket_allowance:int = 2000

emp_dict:dict =  special_emp[emp_id]

insurance_total = emp_dict['insurance']['cat']

