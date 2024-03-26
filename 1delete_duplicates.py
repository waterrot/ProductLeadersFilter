import pandas as pd

# Load the Excel file into a DataFrame
file_path = 'input.xlsx'
df = pd.read_excel(file_path)

# Remove duplicate rows based on 'naam' and 'company' columns
df.drop_duplicates(subset=['Naam', 'Company'], inplace=True)

# Write the updated DataFrame back to Excel
output_file_path = '6trash_leads_no_duplicates.xlsx'
df.to_excel(output_file_path, index=False)

print("Duplicate rows based on 'naam' and 'company' columns removed and saved to", output_file_path)
