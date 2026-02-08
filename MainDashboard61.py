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
ema_cross_analysis = [0, 1, 1, 0]
btc_funding_rates = [0.01, -0.005, 0.002, 0.003]
btc_etf_inflows = [100, 200, 150, 300]
cbbc_components = [0.5, 0.7, 0.65, 0.6]

# Adding headers
headers = ['Date', 'Risk Signal', 'USDT Dominance', 'Fear & Greed Index', 'Golden Cross', 'EMA Cross', 'BTC Funding Rates', 'BTC ETF Inflows', 'CBBI']
worksheet.append(headers)

# Adding sample data
for i in range(len(risk_signals)):
    worksheet.append([
        '2026-02-08', 
        risk_signals[i], 
        usdt_dominance[i], 
        fear_greed_index[i], 
        golden_cross[i], 
        ema_cross_analysis[i], 
        btc_funding_rates[i], 
        btc_etf_inflows[i], 
        cbbc_components[i]
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