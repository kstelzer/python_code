import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

yesterday = pd.read_excel("C:\\Users\\kimbe\\code\\r\\LIST_data_12002.xlsx",na_values=np.nan,header=None)
today = pd.read_excel("C:\\Users\\kimbe\\code\\r\\LIST_data_31646.xlsx",na_values=np.nan,header=None)


rt,ct = yesterday.shape
rtest,ctest = today.shape

df = pd.DataFrame(columns=['Cell_Location','Yesterday_Value','Today_Value'])

for rowNo in range(max(rt,rtest)):
  for colNo in range(max(ct,ctest)):
    # Fetching the template value at a cell
    try:
        yesterday_val = yesterday.iloc[rowNo,colNo]
    except:
        yesterday_val = np.nan

    # Fetching the testsheet value at a cell
    try:
        today_val = today.iloc[rowNo,colNo]
    except:
        today_val = np.nan

    # Comparing the values
    if (str(yesterday_val)!=str(today_val)):
        cell = xl_rowcol_to_cell(rowNo, colNo)
        dfTemp = pd.DataFrame([[cell,yesterday_val,today_val]],
                              columns=['Cell_Location','Yesterday_Value','Today_Value'])
        df = df.append(dfTemp)

df.to_excel("C:\\Users\\kimbe\\code\\r\\diff.xlsx")