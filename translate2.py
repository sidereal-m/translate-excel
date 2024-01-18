import pandas as pd
from googletrans import Translator 
import time
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog

def translate_text(text, source_lang, target_lang):
    if pd.isna(text):  # Skip translation for blank cells
        return ''
    
    try:
        translated = translator.translate(text, src=source_lang, dest=target_lang)
        time.sleep(0.5)
        return text + ' ' + translated.text  # Append translation to existing text
    except Exception as e:
        print(f"Translation failed for text: {text}. Error: {e}")
        return text

def translate_and_save(file_path, source_lang, target_lang):
    if not file_path:
        messagebox.showerror("File Not Selected", "Please select an Excel file.")
        return
    
    # Read all sheets in the Excel file
    xls = pd.ExcelFile(file_path)
    
    translated_dfs = {}

    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)

        # Use map to apply translation to each cell in the DataFrame
        df_translated = df.apply(lambda col: col.map(lambda x: translate_text(x, source_lang, target_lang)))

        translated_dfs[sheet_name] = df_translated

    # Save the translated DataFrames to a new Excel file
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_path:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            for sheet_name, translated_df in translated_dfs.items():
                translated_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    messagebox.showinfo("Translation Complete", "Translation and files saved successfully!")

def select_target_file():
    target_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if target_file:
        target_file_path.delete(0, tk.END)
        target_file_path.insert(0, target_file)
        messagebox.showinfo("Selected File", f"Selected file: {target_file}")

# Language codes and their full names
lang_dict = {
    'Afrikaans': 'af', 'Albanian': 'sq', 'Amharic': 'am', 'Arabic': 'ar', 'Armenian': 'hy', 'Azerbaijani': 'az',
    'Basque': 'eu', 'Belarusian': 'be', 'Bengali': 'bn', 'Bosnian': 'bs', 'Bulgarian': 'bg', 'Catalan': 'ca',
    'Cebuano': 'ceb', 'Chichewa': 'ny', 'Chinese (Simplified)': 'zh-cn', 'Chinese (Traditional)': 'zh-tw',
    'Corsican': 'co', 'Croatian': 'hr', 'Czech': 'cs', 'Danish': 'da', 'Dutch': 'nl', 'English': 'en', 'Esperanto': 'eo',
    'Estonian': 'et', 'Filipino': 'tl', 'Finnish': 'fi', 'French': 'fr', 'Frisian': 'fy', 'Galician': 'gl',
    'Georgian': 'ka', 'German': 'de', 'Greek': 'el', 'Gujarati': 'gu', 'Haitian Creole': 'ht', 'Hausa': 'ha',
    'Hawaiian': 'haw', 'Hebrew': 'he', 'Hindi': 'hi', 'Hmong': 'hmn', 'Hungarian': 'hu', 'Icelandic': 'is',
    'Igbo': 'ig', 'Indonesian': 'id', 'Irish': 'ga', 'Italian': 'it', 'Japanese': 'ja', 'Javanese': 'jv',
    'Kannada': 'kn', 'Kazakh': 'kk', 'Khmer': 'km', 'Kinyarwanda': 'rw', 'Korean': 'ko', 'Kurdish (Kurmanji)': 'ku',
    'Kyrgyz': 'ky', 'Lao': 'lo', 'Latin': 'la', 'Latvian': 'lv', 'Lithuanian': 'lt', 'Luxembourgish': 'lb',
    'Macedonian': 'mk', 'Malagasy': 'mg', 'Malay': 'ms', 'Malayalam': 'ml', 'Maltese': 'mt', 'Maori': 'mi',
    'Marathi': 'mr', 'Mongolian': 'mn', 'Burmese': 'my', 'Nepali': 'ne', 'Norwegian': 'no', 'Odia': 'or',
    'Pashto': 'ps', 'Persian': 'fa', 'Polish': 'pl', 'Portuguese': 'pt', 'Punjabi': 'pa', 'Romanian': 'ro',
    'Russian': 'ru', 'Samoan': 'sm', 'Scots Gaelic': 'gd', 'Serbian': 'sr', 'Sesotho': 'st', 'Shona': 'sn',
    'Sindhi': 'sd', 'Sinhala': 'si', 'Slovak': 'sk', 'Slovenian': 'sl', 'Somali': 'so', 'Spanish': 'es',
    'Sundanese': 'su', 'Swahili': 'sw', 'Swedish': 'sv', 'Tajik': 'tg', 'Tamil': 'ta', 'Tatar': 'tt', 'Telugu': 'te',
    'Thai': 'th', 'Turkish': 'tr', 'Turkmen': 'tk', 'Ukrainian': 'uk', 'Urdu': 'ur', 'Uyghur': 'ug', 'Uzbek': 'uz',
    'Vietnamese': 'vi', 'Welsh': 'cy', 'Xhosa': 'xh', 'Yiddish': 'yi', 'Yoruba': 'yo', 'Zulu': 'zu'
}

# UI setup
root = tk.Tk()
root.title("Excel Translator")
root.geometry("400x350")  # Set initial width and height

translator = Translator()

header_frame = tk.Frame(root, bg="light blue", padx=10, pady=10)
header_frame.grid(row=0, column=0, columnspan=2, sticky="ew")

header_label = tk.Label(header_frame, text="Excel Translator", font=("Arial", 14), bg="light blue")
header_label.pack()

target_file_frame = tk.Frame(root)
target_file_frame.grid(row=1, column=0, columnspan=2, pady=5)

target_file_button = tk.Button(target_file_frame, text="Select Target File", command=select_target_file)
target_file_button.pack(side=tk.LEFT, padx=5)

target_file_path = tk.Entry(target_file_frame, width=40)
target_file_path.pack(side=tk.LEFT)

source_lang_label = tk.Label(root, text="Select Source Language:")
source_lang_label.grid(row=3, column=0, padx=10, sticky="w")

var_source = tk.StringVar(root)
var_source.set("German")  # Default source language: English

options_source = list(lang_dict.keys())
dropdown_source = tk.OptionMenu(root, var_source, *options_source)
dropdown_source.grid(row=3, column=1, sticky="ew")

language_label = tk.Label(root, text="Select Target Language:")
language_label.grid(row=4, column=0, padx=10, sticky="w")

var_target = tk.StringVar(root)
var_target.set("English")  # Default target language: German

options_target = list(lang_dict.keys())
dropdown_target = tk.OptionMenu(root, var_target, *options_target)
dropdown_target.grid(row=4, column=1, sticky="ew")

button = tk.Button(root, text="Translate", command=lambda: translate_and_save(target_file_path.get(), var_source.get(), var_target.get()))
button.grid(row=5, column=0, columnspan=2, pady=10)

root.mainloop()
