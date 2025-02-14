import pandas as pd


text = pd.read_excel('tm2025-sm.xlsx', sheet_name='Лист1')

print(text.head())