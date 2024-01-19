import os
import pandas as pd
from googletrans import Translator


class ExcelTranslatorLogic:
    def __init__(self):
        self.translator = Translator()

    def translate_text(self, text, source_lang, target_lang):
        if pd.isna(text):  # Skip translation for blank cells
            return ''

        try:
            translated = self.translator.translate(text, src=source_lang, dest=target_lang)
            return text + ' ' + translated.text  # Append translation to existing text
        except Exception as e:
            print(f"Translation failed for text: {text}. Error: {e}")
            return text

    def translate_and_save(self, file_path, source_lang, target_lang):
        if not file_path:
            print("Please provide a valid file path.")
            return

        # Read all sheets in the Excel file
        xls = pd.ExcelFile(file_path)

        translated_dfs = {}

        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)

            # Use map to apply translation to each cell in the DataFrame
            df_translated = df.apply(lambda col: col.map(lambda x: self.translate_text(x, source_lang, target_lang)))

            translated_dfs[sheet_name] = df_translated

        # Create the directory if it doesn't exist
        output_directory = os.path.join(os.path.dirname(file_path), "translated")
        os.makedirs(output_directory, exist_ok=True)

        # Save the translated DataFrames to a new Excel file
        output_path = os.path.join(output_directory, f"translated_{os.path.basename(file_path)}")
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            for sheet_name, translated_df in translated_dfs.items():
                translated_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Translation complete. Output saved to: {output_path}")
