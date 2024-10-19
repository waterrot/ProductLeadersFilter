import pandas as pd

# Lees het Excel-bestand met resultaten in
results_df = pd.read_excel("4leadlist_with_language_column.xlsx")

# Filter de rijen waar de waarde in de kolom "Nederlandse Naam" gelijk is aan 1
nederlandse_leads_df = results_df[results_df['Nederlandse Naam'] == 1].copy()

# Verwijder de kolom "Nederlandse Naam" uit de dataframe van Nederlandse leads
nederlandse_leads_df.drop(columns=['Nederlandse Naam'], inplace=True)

# Schrijf de Nederlandse leads naar een nieuw Excel-bestand
nederlandse_leads_df.to_excel("5nederlandse_leads.xlsx", index=False)

# Filter de rijen waar de waarde in de kolom "Nederlandse Naam" niet gelijk is aan 1
overige_leads_df = results_df[results_df['Nederlandse Naam'] != 1]

# Schrijf de overige leads naar het bestand "engelse_leads.xlsx"
overige_leads_df.to_excel("5engelse_leads.xlsx", index=False)
