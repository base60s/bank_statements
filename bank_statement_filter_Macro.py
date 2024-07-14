import pandas as pd
from datetime import datetime
import calendar

def get_week_number(date):
    return date.isocalendar()[1]

def process_bank_statement(input_file, output_file, bank_name):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Process the data
    result = []
    for _, row in df.iterrows():
        date = row['Fecha de imputación']
        concept = row['Descripción']
        amount = row['Importe']

        # Format date
        formatted_date = date.strftime('%d/%m/%Y')

        # Calculate week number
        week_number = get_week_number(date)

        # Determine if it's a debit or credit
        debit = abs(amount) if amount < 0 else 0
        credit = amount if amount > 0 else 0

        result.append([formatted_date, week_number, bank_name, concept, debit, credit])

    # Create output DataFrame
    output_df = pd.DataFrame(result, columns=['Fecha', '# Semana', 'Banco', 'Concepto', 'Debitos', 'Creditos'])

    # Save to CSV
    output_df.to_csv(output_file, index=False)

    print(f"Processing complete. Output saved to {output_file}")

if __name__ == "__main__":
    input_file = "Macro.xlsx"
    output_file = "output.csv"
    bank_name = "Macro"  # You can change this to the appropriate bank name
    process_bank_statement(input_file, output_file, bank_name)