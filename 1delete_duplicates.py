import pandas as pd

# Load the Excel file into a DataFrame
file_path = 'input.xlsx'
df = pd.read_excel(file_path)

# Get the column indices for name, job title, and company from the user
print("Please enter the column indices for the following fields:")
print("Remember that Python indices start with 0, so A=0, B=1, C=2....")

# Function to get a valid index from the user
def get_column_index(column_name):
    while True:
        try:
            index = int(input(f"Enter the index for '{column_name}': "))
            if index < 0 or index >= len(df.columns):
                print("Index is out of range. Please enter a valid index.")
            else:
                return index
        except ValueError:
            print("Invalid input. Please enter an integer.")

# Get indices for the required columns
name_index = get_column_index("Name")
title_index = get_column_index("Title")
organisation_index = get_column_index("Company")

# Remove duplicate rows based on name, company, and title columns
df.drop_duplicates(subset=[df.columns[name_index], df.columns[organisation_index], df.columns[title_index]], inplace=True)

# Write the updated DataFrame back to Excel
output_file_path = '1trash_leads_no_duplicates.xlsx'
df.to_excel(output_file_path, index=False)

print("Duplicate rows based on 'name', 'company', and 'title' columns removed and saved to", output_file_path)
