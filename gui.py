from PyQt5.QtWidgets import QMainWindow, QLabel, QPushButton, QFileDialog, QComboBox, QVBoxLayout, QWidget, QApplication, QMessageBox
from logic import ExcelTranslatorLogic


class ExcelTranslatorGUI(QMainWindow):
    def __init__(self, logic):
        super().__init__()
        self.logic = logic

        self.setWindowTitle("Excel Translator")
        self.setGeometry(100, 100, 400, 300)

        self.setup_ui()

    def setup_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        self.label_header = QLabel("Excel Translator", self)
        self.label_header.setStyleSheet("font-size: 14pt; background-color: lightblue;")
        layout.addWidget(self.label_header)

        self.label_file = QLabel("Select Target File:", self)
        layout.addWidget(self.label_file)

        self.file_path = QLabel("", self)
        layout.addWidget(self.file_path)

        self.btn_select_file = QPushButton("Select File", self)
        self.btn_select_file.clicked.connect(self.select_file)
        layout.addWidget(self.btn_select_file)

        self.label_source_lang = QLabel("Select Source Language:", self)
        layout.addWidget(self.label_source_lang)

        self.combo_source_lang = QComboBox(self)
        layout.addWidget(self.combo_source_lang)
        self.populate_language_combobox(self.combo_source_lang)

        self.label_target_lang = QLabel("Select Target Language:", self)
        layout.addWidget(self.label_target_lang)

        self.combo_target_lang = QComboBox(self)
        layout.addWidget(self.combo_target_lang)
        self.populate_language_combobox(self.combo_target_lang)

        self.btn_translate = QPushButton("Translate", self)
        self.btn_translate.clicked.connect(self.translate_and_save)
        layout.addWidget(self.btn_translate)

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

    def translate_and_save(self):
        file_path = self.file_path.text()
        source_lang = self.combo_source_lang.currentText()
        target_lang = self.combo_target_lang.currentText()

        self.logic.translate_and_save(file_path, source_lang, target_lang)
        QMessageBox.information(self, "Translation Complete", "Translation and files saved successfully!")


if __name__ == '__main__':
    app = QApplication([])
    logic = ExcelTranslatorLogic()
    gui = ExcelTranslatorGUI(logic)
    gui.show()
    app.exec_()
