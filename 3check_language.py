import pandas as pd
from openai import OpenAI

# Vul je OpenAI API-sleutel in
api_key = OpenAI(api_key="")

# De datafile
file_path = '8leads_without_trash.xlsx'


class SortNames:
    def __init__(self, api_key, file_path):
        # Set the OpenAI API key
        self.api_key = api_key
        self.file_path = file_path

    def is_nederlandse_naam(self, naam):
        # Prepare a prompt for ChatGPT
        prompt = f"Is {naam} een nederlandse naam? Bij twijfel is de achternaam leidend. Geef \
            als output echt alleen een 1 weer als het een nederlandse naam is en anders een 0. geen \
                uitleg erbij of andere tekst naast de 0 of 1"
        try:
            # Send the prompt to ChatGPT
            response = self.api_key.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=150
            )

            # Extract and return the generated response
            content = response.choices[0].message.content
            # Extracting the last character from the content
            result = int(content.strip()[-1])
            return result

        except Exception as e:
            print(f"Error during API request: {e}")
            return None

    # Functie om de Excel-bestanden te verwerken
    def process_excel(self):
        # Lees het Excel-bestand
        df = pd.read_excel(self.file_path)
        
        # Voeg een nieuwe kolom toe
        df['Nederlandse Naam'] = df['Naam'].apply(lambda x: self.is_nederlandse_naam(x))

        # Schrijf het resultaat terug naar het Excel-bestand
        df.to_excel("9leadlist_with_language_column.xlsx", index=False)

# Create an instance of the NameProcessor class
name_processor = SortNames(api_key, file_path)

# Call the process_excel method to process the Excel file
name_processor.process_excel()
