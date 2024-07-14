import pandas as pd
from datetime import datetime
import csv

def get_week_number(date_str):
    date = datetime.strptime(date_str, '%d/%m/%Y')
    return date.isocalendar()[1]

def process_bank_statement(input_file, output_file, bank_name):
    # Read the Excel file, skipping the first 5 rows
    df = pd.read_excel(input_file, skiprows=5)

    # Initialize lists to store the processed data
    dates = []
    week_numbers = []
    bank_names = []
    concepts = []
    debits = []
    credits = []

    # Process each row in the dataframe
    for _, row in df.iterrows():
        date = row['Fecha']
        concept = row['Concepto']
        amount_str = str(row['Importe']).strip()  # Remove leading/trailing whitespace

        # Ensure date is in the correct format
        if isinstance(date, str):
            formatted_date = date
        else:
            formatted_date = date.strftime('%d/%m/%Y')

        # Calculate week number
        week_number = get_week_number(formatted_date)

        # Remove "$" symbol and categorize amount as debit or credit based on the presence of "-"
        amount_str = amount_str.replace('$', '').strip()
        if '-' in amount_str:
            debit = amount_str.replace('-', '').strip()  # Remove the "-" and any surrounding whitespace
            credit = ''
        else:
            debit = ''
            credit = amount_str

        # Append data to lists
        dates.append(formatted_date)
        week_numbers.append(week_number)
        bank_names.append(bank_name)
        concepts.append(concept)
        debits.append(debit)
        credits.append(credit)

    # Write the processed data to a CSV file
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Fecha', '# Semana', 'Banco', 'Concepto', 'Debitos', 'Creditos'])
        for i in range(len(dates)):
            writer.writerow([dates[i], week_numbers[i], bank_names[i], concepts[i], debits[i], credits[i]])

    print(f"Processed data has been written to {output_file}")

# Usage
input_file = 'Nacion.xlsx'
output_file = 'output.csv'
bank_name = 'Nacion'
process_bank_statement(input_file, output_file, bank_name)