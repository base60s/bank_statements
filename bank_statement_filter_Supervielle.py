import pandas as pd
from datetime import datetime

def get_week_number(date):
    return date.isocalendar()[1]

# Read the Excel file
df = pd.read_excel('Supervielle.xlsx')

# Convert 'Fecha' to datetime and format it
df['Fecha'] = pd.to_datetime(df['Fecha'])
df['Fecha'] = df['Fecha'].dt.strftime('%d/%m/%Y')

# Calculate week number
df['# Semana'] = df['Fecha'].apply(lambda x: get_week_number(datetime.strptime(x, '%d/%m/%Y')))

# Add bank name column
df['Banco'] = 'Supervielle'

# Rename columns
df = df.rename(columns={'Débito': 'Debitos', 'Crédito': 'Creditos'})

# Select and reorder columns
output_df = df[['Fecha', '# Semana', 'Banco', 'Concepto', 'Debitos', 'Creditos']]

# Write to CSV
output_df.to_csv('output.csv', index=False)

print("Processing complete. Output file 'output.csv' has been created.")