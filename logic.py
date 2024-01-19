import os
import pandas as pd
from googletrans import Translator
import configparser


class ExcelTranslatorLogic:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super(ExcelTranslatorLogic, cls).__new__(cls, *args, **kwargs)
            cls._instance.translator = Translator()
            cls._instance.load_config()
        return cls._instance

    def load_config(self):
        self.config = configparser.ConfigParser()
        self.config_file_path = os.path.join(os.path.dirname(__file__), 'config.ini')
        self.config.read(self.config_file_path)

        if not self.config.has_section('DefaultLanguages'):
            self.config.add_section('DefaultLanguages')

        self.default_source_lang = self.config.get('DefaultLanguages', 'DefaultSourceLang', fallback="en")
        self.default_target_lang = self.config.get('DefaultLanguages', 'DefaultTargetLang', fallback="fr")

    def save_config(self):
        if not self.config.has_section('DefaultLanguages'):
            self.config.add_section('DefaultLanguages')

        self.config.set('DefaultLanguages', 'DefaultSourceLang', self.default_source_lang)
        self.config.set('DefaultLanguages', 'DefaultTargetLang', self.default_target_lang)

        with open(self.config_file_path, 'w') as configfile:
            self.config.write(configfile)

    def translate_text(self, text, source_lang=None, target_lang=None):
        if pd.isna(text):
            return ''

        source_lang = source_lang or self.default_source_lang
        target_lang = target_lang or self.default_target_lang

        try:
            translated = self.translator.translate(text, src=source_lang, dest=target_lang)
            return text + ' ' + translated.text
        except Exception as e:
            print(f"Translation failed for text: {text}. Error: {e}")
            return text

    def translate_and_save(self, file_path, source_lang=None, target_lang=None):
        if not file_path:
            print("Please provide a valid file path.")
            return

        xls = pd.ExcelFile(file_path)
        translated_dfs = {}

        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)
            df_translated = df.apply(lambda col: col.map(lambda x: self.translate_text(x, source_lang, target_lang)))
            translated_dfs[sheet_name] = df_translated

        output_directory = os.path.join(os.path.dirname(file_path), "translated")
        os.makedirs(output_directory, exist_ok=True)

        output_path = os.path.join(output_directory, f"translated_{os.path.basename(file_path)}")
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            for sheet_name, translated_df in translated_dfs.items():
                translated_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Translation complete. Output saved to: {output_path}")

    def set_default_languages(self, default_source_lang, default_target_lang):
        self.default_source_lang = default_source_lang
        self.default_target_lang = default_target_lang
        self.save_config()
