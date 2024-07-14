import pandas as pd
from datetime import datetime
import re

def read_file(file_path):
    if file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_path}")

def clean_monetary_value(value):
    if isinstance(value, str):
        cleaned = re.sub(r'[^\d.-]', '', value)
        return float(cleaned) if cleaned else 0.0
    return float(value) if pd.notnull(value) else 0.0

def process_bank_statement(input_file, output_file):
    try:
        # Read the file
        df = read_file(input_file)

        # Print column names and first few rows for debugging
        print("Column names:", df.columns)
        print("\nFirst few rows of the original dataframe:")
        print(df.head())

        # Convert APPROVAL_DATE to datetime for week calculation
        df['APPROVAL_DATE'] = pd.to_datetime(df['APPROVAL_DATE'])

        # Extract date and calculate week number
        df['Fecha'] = df['APPROVAL_DATE'].dt.strftime('%d/%m/%Y')
        df['# Semana'] = df['APPROVAL_DATE'].dt.isocalendar().week

        # Clean and process SETTLEMENT_NET_AMOUNT
        df['SETTLEMENT_NET_AMOUNT'] = df['SETTLEMENT_NET_AMOUNT'].apply(clean_monetary_value)

        # Determine Debitos and Creditos based on SETTLEMENT_NET_AMOUNT
        df['Debitos'] = df['SETTLEMENT_NET_AMOUNT'].apply(lambda x: abs(x) if x < 0 else 0)
        df['Creditos'] = df['SETTLEMENT_NET_AMOUNT'].apply(lambda x: x if x >= 0 else 0)

        # Create Concepto from available information
        df['Concepto'] = df['TRANSACTION_TYPE'] + ' - ' + df['PAYMENT_METHOD'].fillna('') + ' - ' + df['EXTERNAL_REFERENCE'].fillna('')

        # Create a new 'Banco' column with 'MercadoPago' for all rows
        df['Banco'] = 'MercadoPago'

        # Select and reorder columns
        result = df[['Fecha', '# Semana', 'Banco', 'Concepto', 'Debitos', 'Creditos']]

        # Write to CSV
        result.to_csv(output_file, index=False, encoding='utf-8')

        print(f"\nProcessed data has been written to {output_file}")
        
        # Print the first few rows of the result for verification
        print("\nFirst few rows of the processed dataframe:")
        print(result.head())
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        print("Please make sure the input file exists and is in the correct format.")

if __name__ == "__main__":
    input_file = "MercadoPago.xlsx"
    output_file = "output.csv"
    process_bank_statement(input_file, output_file)
