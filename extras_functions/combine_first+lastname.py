import pandas as pd

# Laad het CSV-bestand
input_file = 'Order FO831A6421F08 status.csv'
df = pd.read_csv(input_file)

# Voeg de "Person - Name" kolom toe als laatste kolom
df['Person - Name'] = df['First Name'] + ' ' + df['Last Name']

# Sla het gewijzigde CSV-bestand op
output_file = 'output.csv'
df.to_csv(output_file, index=False)
