import pandas as pd

# Read the Excel file
df = pd.read_excel('../datasheets/Imports/Urkunden_Metadatenliste_v11.xlsx')

# Initialize the new column
df['P1069 (administrative Zugehörigkeit)'] = ''

# Iterate through each row and check for patterns in 'Zusammenfassung (P724)'
for index, row in df.iterrows():
    summary = str(row['Zusammenfassung (P724) '])
    if '(TH)' in summary:
        df.at[index, 'P1069 (administrative Zugehörigkeit)'] = 'Q1306300'
    elif '(TS)' in summary:
        df.at[index, 'P1069 (administrative Zugehörigkeit)'] = 'Q1306302'
    elif '(TL)' in summary:
        df.at[index, 'P1069 (administrative Zugehörigkeit)'] = 'Q1306299'
    elif '(TP)' in summary:
        df.at[index, 'P1069 (administrative Zugehörigkeit)'] = 'Q1306301'

# Save the updated dataframe as v12
df.to_excel('../datasheets/Imports/Urkunden_Metadatenliste_v12.xlsx', index=False)