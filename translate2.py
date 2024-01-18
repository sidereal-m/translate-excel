import sys
import pandas as pd
from googletrans import Translator
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QPushButton, QFileDialog, QComboBox

class ExcelTranslator(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Translator")
        self.setGeometry(100, 100, 400, 300)

        self.translator = Translator()

        self.setup_ui()

    def setup_ui(self):
        self.label_header = QLabel("Excel Translator", self)
        self.label_header.setGeometry(10, 10, 380, 30)
        self.label_header.setStyleSheet("font-size: 14pt; background-color: lightblue;")

        self.label_file = QLabel("Select Target File:", self)
        self.label_file.setGeometry(10, 50, 150, 20)

        self.file_path = QLabel("", self)
        self.file_path.setGeometry(160, 50, 180, 20)

        self.btn_select_file = QPushButton("Select File", self)
        self.btn_select_file.setGeometry(350, 50, 40, 20)
        self.btn_select_file.clicked.connect(self.select_file)

        self.label_source_lang = QLabel("Select Source Language:", self)
        self.label_source_lang.setGeometry(10, 80, 150, 20)

        self.combo_source_lang = QComboBox(self)
        self.combo_source_lang.setGeometry(160, 80, 180, 20)
        self.populate_language_combobox(self.combo_source_lang)

        self.label_target_lang = QLabel("Select Target Language:", self)
        self.label_target_lang.setGeometry(10, 110, 150, 20)

        self.combo_target_lang = QComboBox(self)
        self.combo_target_lang.setGeometry(160, 110, 180, 20)
        self.populate_language_combobox(self.combo_target_lang)

        self.btn_translate = QPushButton("Translate", self)
        self.btn_translate.setGeometry(10, 150, 380, 30)
        self.btn_translate.clicked.connect(self.translate_and_save)

    def populate_language_combobox(self, combo_box):
        languages = [
            'Afrikaans', 'Albanian', 'Amharic', 'Arabic', 'Armenian', 'Azerbaijani', 'Basque', 'Belarusian', 'Bengali',
            'Bosnian', 'Bulgarian', 'Catalan', 'Cebuano', 'Chichewa', 'Chinese (Simplified)', 'Chinese (Traditional)',
            'Corsican', 'Croatian', 'Czech', 'Danish', 'Dutch', 'English', 'Esperanto', 'Estonian', 'Filipino', 'Finnish',
            'French', 'Frisian', 'Galician', 'Georgian', 'German', 'Greek', 'Gujarati', 'Haitian Creole', 'Hausa', 'Hawaiian',
            'Hebrew', 'Hindi', 'Hmong', 'Hungarian', 'Icelandic', 'Igbo', 'Indonesian', 'Irish', 'Italian', 'Japanese',
            'Javanese', 'Kannada', 'Kazakh', 'Khmer', 'Kinyarwanda', 'Korean', 'Kurdish (Kurmanji)', 'Kyrgyz', 'Lao', 'Latin',
            'Latvian', 'Lithuanian', 'Luxembourgish', 'Macedonian', 'Malagasy', 'Malay', 'Malayalam', 'Maltese', 'Maori',
            'Marathi', 'Mongolian', 'Burmese', 'Nepali', 'Norwegian', 'Odia', 'Pashto', 'Persian', 'Polish', 'Portuguese',
            'Punjabi', 'Romanian', 'Russian', 'Samoan', 'Scots Gaelic', 'Serbian', 'Sesotho', 'Shona', 'Sindhi', 'Sinhala',
            'Slovak', 'Slovenian', 'Somali', 'Spanish', 'Sundanese', 'Swahili', 'Swedish', 'Tajik', 'Tamil', 'Tatar', 'Telugu',
            'Thai', 'Turkish', 'Turkmen', 'Ukrainian', 'Urdu', 'Uyghur', 'Uzbek', 'Vietnamese', 'Welsh', 'Xhosa', 'Yiddish',
            'Yoruba', 'Zulu'
        ]

        combo_box.addItems(languages)

    def select_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx;*.xls)")
        if file_path:
            self.file_path.setText(file_path)

    def translate_text(self, text, source_lang, target_lang):
        if pd.isna(text):  # Skip translation for blank cells
            return ''

        try:
            translated = self.translator.translate(text, src=source_lang, dest=target_lang)
            return text + ' ' + translated.text  # Append translation to existing text
        except Exception as e:
            print(f"Translation failed for text: {text}. Error: {e}")
            return text

    def translate_and_save(self):
        file_path = self.file_path.text()
        source_lang = self.combo_source_lang.currentText()
        target_lang = self.combo_target_lang.currentText()

        if not file_path:
            QMessageBox.warning(self, "File Not Selected", "Please select an Excel file.")
            return

        # Read all sheets in the Excel file
        xls = pd.ExcelFile(file_path)

        translated_dfs = {}

        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)

            # Use map to apply translation to each cell in the DataFrame
            df_translated = df.apply(lambda col: col.map(lambda x: self.translate_text(x, source_lang, target_lang)))

            translated_dfs[sheet_name] = df_translated

        # Save the translated DataFrames to a new Excel file
        output_path, _ = QFileDialog.getSaveFileName(self, "Save Translated Excel File", "", "Excel Files (*.xlsx)")
        if output_path:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                for sheet_name, translated_df in translated_dfs.items():
                    translated_df.to_excel(writer, sheet_name=sheet_name, index=False)

            QMessageBox.information(self, "Translation Complete", "Translation and files saved successfully!")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelTranslator()
    window.show()
    sys.exit(app.exec_())
