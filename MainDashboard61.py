import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook

# Create a workbook and add a worksheet
workbook = Workbook()
worksheet = workbook.active
worksheet.title = 'Dashboard'

# Sample data for the dashboard
risk_signals = ['Risk On', 'Risk Off', 'Risk On', 'Risk Off']
usdt_dominance = [0.45, 0.47, 0.44, 0.49]
fear_greed_index = [50, 60, 40, 35]
golden_cross = [10000, 12000, 11000, 11500]

# Adding sample data
for i in range(len(risk_signals)):
    worksheet.append([
        '2026-02-08', 
        risk_signals[i], 
        usdt_dominance[i], 
        fear_greed_index[i], 
        golden_cross[i]
    ])

# Save the workbook
workbook.save('Macroeconomics_Dashboard.xlsx')

# Visualization (not stored in Excel)
plt.figure(figsize=(10, 6))
plt.plot(usdt_dominance, label='USDT Dominance', marker='o')
plt.title('USDT Dominance Over Time')
plt.xlabel('Time')
plt.ylabel('USDT Dominance')
plt.legend()
plt.grid()
plt.show()