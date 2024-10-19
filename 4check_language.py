import pandas as pd
from openai import OpenAI

# Vul je OpenAI API-sleutel in
api_key = OpenAI(api_key="ENTER VALUE HERE")

# De datafile
file_path = '3leads_sorted.xlsx'


class SortNames:
    def __init__(self, api_key, file_path):
        # Set the OpenAI API key
        self.api_key = api_key
        self.file_path = file_path

    def is_nederlandse_naam(self, naam):
        # Prepare a prompt for ChatGPT
        prompt = f"Is {naam} een nederlandse naam? Bij twijfel is de achternaam leidend en doorslaggevend! Geef een 1 weer als het een nederlandse naam is en anders een 0. geen uitleg erbij of andere tekst naast de 0 of 1 (heb alleen de cijfers nodig, als er tekst bijkomt werkt het niet meer)."
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
            
            # Debugging: Print the content to see the exact response
            print(f"API Response for {naam}: '{content}'")
            
            # Check if the response is exactly '1' or '0'
            if result == 1:
                return 1
            elif result == 0:
                return 0
            else:
                print(f"Unexpected response content for {naam}: '{content}'")
                return None

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
        df.to_excel("4leadlist_with_language_column.xlsx", index=False)

# Create an instance of the NameProcessor class
name_processor = SortNames(api_key, file_path)

# Call the process_excel method to process the Excel file
name_processor.process_excel()
