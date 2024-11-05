import pandas as pd
from datetime import datetime,timedelta
import numpy as np
from dateutil.relativedelta import relativedelta
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt,RGBColor, Cm, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import calendar

data = {'invoice_number':['NBL/IVL241767','NBL/IVL241767','NBL/IVL241767','NBL/IVL241768','NBL/IVL241768','NBL/IVL241768'],'amt':[1000,1000,1000,2000,2000,2000]}

df = pd.DataFrame(data=data).drop_duplicates(keep='first',ignore_index=True)['amt'].sum()

# df.drop_duplicates(keep='first',inplace=True)

