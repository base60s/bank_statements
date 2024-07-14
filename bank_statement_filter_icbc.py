import pandas as pd
from datetime import datetime
import sys

def get_week_number(date):
    return date.isocalendar()[1]

def remove_negative_sign(value):
    if isinstance(value, (int, float)):
        return abs(value)
    elif isinstance(value, str):
        return value.lstrip('-')
    return value

try:
    # Read the Excel file
    df = pd.read_excel('icbc.xlsx', skiprows=1)

    # Print column names for debugging
    print("Original column names:", df.columns.tolist())

    # Rename columns, handling potential variations in spelling
    column_mapping = {
        'Fecha contable': 'Fecha',
        'Concepto': 'Concepto',
        'Debito en $': 'Debitos',
    }

    # Find the credit column (it might be 'Créditos' or 'Creditos')
    credit_column = [col for col in df.columns if 'cr' in col.lower()]
    if credit_column:
        column_mapping[credit_column[0]] = 'Creditos'
    else:
        raise ValueError("Credit column not found in the Excel file")

    df = df.rename(columns=column_mapping)

    # Print column names after renaming for debugging
    print("Column names after renaming:", df.columns.tolist())

    # Convert 'Fecha' to datetime and format it
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y', errors='coerce')
    df['Fecha'] = df['Fecha'].dt.strftime('%d/%m/%Y')

    # Calculate week number
    df['# Semana'] = df['Fecha'].apply(lambda x: get_week_number(datetime.strptime(x, '%d/%m/%Y')) if pd.notna(x) else None)

    # Add 'Banco' column with 'ICBC' as the value
    df['Banco'] = 'ICBC'

    # Remove negative signs from 'Debitos' and 'Creditos' columns
    df['Debitos'] = df['Debitos'].apply(remove_negative_sign)
    df['Creditos'] = df['Creditos'].apply(remove_negative_sign)

    # Select and reorder columns
    output_columns = ['Fecha', '# Semana', 'Banco', 'Concepto', 'Debitos', 'Creditos']
    output_df = df[output_columns]

    # Save to CSV
    output_df.to_csv('output.csv', index=False)

    print("Processing complete. Output saved to 'output.csv'.")

except FileNotFoundError:
    print("Error: The file 'icbc.xlsx' was not found. Please make sure it's in the same directory as the script.")
except ValueError as e:
    print(f"Error: {str(e)}")
except KeyError as e:
    print(f"Error: Column {str(e)} not found in the Excel file. Please check the column names.")
except Exception as e:
    print(f"An unexpected error occurred: {str(e)}")
    print("Please check your input file and try again.")