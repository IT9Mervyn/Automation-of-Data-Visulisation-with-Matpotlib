import csv
from turtle import title
import pandas as pd
from pandas import read_excel
import openpyxl

import sys
import matplotlib
matplotlib.use('Qt5Agg')
import tkinter
import matplotlib.pyplot as plt
import numpy as np


file ='C:/Users/merl0330/Desktop/Data _visualisation/2021Q4_YTAV_1st_summary(Raw).xlsx'

Terr_code = []
MR_claim = []
PR_claim = []
FILE_claim = []
Total_claim = []
Market_share = []
Market_share_Q3 = []
wb = openpyxl.load_workbook(file)

sheet = wb.get_sheet_by_name('2021Q4')
for row_of_cell in sheet['E2':'E14']:
    for cell_object in row_of_cell:
        Terr_code.append(cell_object.value)

for row_of_cell in sheet['AH2':'AH14']:
    for cell_object in row_of_cell:
        MR_claim.append(cell_object.value)

for row_of_cell in sheet['AI2':'AI14']:
    for cell_object in row_of_cell:
        PR_claim.append(cell_object.value)
for row_of_cell in sheet['W2':'W14']:
    for cell_object in row_of_cell:
        FILE_claim.append(cell_object.value)

# Q3 market share

for row_of_cell in sheet['E20':'E32']:
    for cell_object in row_of_cell:
        Market_share_Q3.append(cell_object.value*100)

i = 0
for x in MR_claim:
    claim_sum = x + PR_claim[i]
    i = i+1
    Total_claim.append(claim_sum)

i = 0
for y in Total_claim:
    m_share = Total_claim[i]/FILE_claim[i] *100 
    i=i+1
    Market_share.append(m_share)




print(Market_share_Q3)

X_axis = np.arange(len(Terr_code))

# Market share - Bar chart
plt.title('Market Share')
plt.xlabel('Territory')
plt.ylabel('Percentage')
plt.bar(X_axis - 0.2, Market_share, 0.4, label = 'Q4')
plt.bar(X_axis+0.2, Market_share_Q3,0.4, label = 'Q3')
plt.xticks(X_axis, Terr_code)
plt.legend()
plt.savefig('Bar_Chart.png')
plt.close()


# Market share - Pie chart
plt.pie(Market_share, labels= Terr_code)
plt.legend()
plt.savefig('Pie_Chart.png')
plt.close()


# line graph to see trend
plt.plot(Market_share, label = 'Q4')
plt.plot(Market_share_Q3,label = 'Q3')
plt.xticks(X_axis, Terr_code)
plt.legend()
plt.savefig('Line_chart.png')

#Two  lines to make our compiler able to draw:
# plt.savefig(sys.stdout.buffer)
# sys.stdout.flush()
