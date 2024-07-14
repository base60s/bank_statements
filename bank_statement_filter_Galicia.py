import pandas as pd
from datetime import datetime
import calendar

# Constant for the bank name
BANK_NAME = "Galicia"

def get_week_number(date):
    return date.isocalendar()[1]

def process_bank_statement(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Print column names
    print("Columns in the Excel file:")
    print(df.columns)

    # Function to find the correct column name
    def find_column(possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        raise KeyError(f"Could not find any of these columns: {possible_names}")

    # Find the correct column names
    fecha_col = find_column(['Fecha', 'fecha', 'FECHA', 'Date'])
    descripcion_col = find_column(['Descripción', 'descripcion', 'DESCRIPCION', 'Description'])
    debitos_col = find_column(['Débitos', 'debitos', 'DEBITOS', 'Debits'])
    creditos_col = find_column(['Créditos', 'creditos', 'CREDITOS', 'Credits'])

    # Process the data
    df['Fecha'] = pd.to_datetime(df[fecha_col])
    df['# Semana'] = df['Fecha'].apply(get_week_number)
    df['Fecha'] = df['Fecha'].dt.strftime('%d/%m/%Y')
    df['Concepto'] = df[descripcion_col]
    df['Debitos'] = df[debitos_col]
    df['Creditos'] = df[creditos_col]
    df['Banco'] = BANK_NAME  # Add the bank name column

    # Select and reorder columns
    result_df = df[['Fecha', '# Semana', 'Banco', 'Concepto', 'Debitos', 'Creditos']]

    # Write to CSV
    result_df.to_csv(output_file, index=False)

    print(f"Processed data has been saved to {output_file}")

if __name__ == "__main__":
    input_file = "Galicia.xlsx"
    output_file = "output.csv"
    process_bank_statement(input_file, output_file)