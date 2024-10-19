import pandas as pd

def compare_excel_files(output_file):
  """Vergelijkt twee specifieke Excel-bestanden op basis van de kolom "Person - Email - Work" en slaat de resultaten op in een nieuw bestand.

  Args:
    output_file: De naam van het nieuwe Excel-bestand.
  """

  file1 = "11engelse_leads.xlsx"
  file2 = "cto-pipedrive.xlsx"
  email_column = "Person - Email - Work"

  # Lees de Excel-bestanden in als pandas DataFrames
  df1 = pd.read_excel(file1)
  df2 = pd.read_excel(file2)

  # Voer een inner join uit op basis van de opgegeven kolom
  merged_df = pd.merge(df1, df2, on=email_column, how='inner')

  # Sla het resultaat op in een nieuw Excel-bestand
  merged_df.to_excel(output_file, index=False)

# Voorbeeld gebruik:
output_file = "overlappende_leads.xlsx"
compare_excel_files(output_file)
