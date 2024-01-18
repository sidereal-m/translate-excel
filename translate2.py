import pandas as pd
from googletrans import Translator 
import time
import tkinter as tk
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import filedialog

def translate_text(text, target_lang):
    try:
        translated = translator.translate(text, src='en', dest=target_lang)
        time.sleep(0.5)
        return translated.text
    except Exception as e:
        print(f"Translation failed for text: {text}. Error: {e}")
        return ''

def translate_and_save(file_path):
    if not file_path:
        messagebox.showerror("File Not Selected", "Please select an Excel file.")
        return
    
    df = pd.read_excel(file_path)
    column_name = entry.get()
    target_lang = var.get()
    
    if column_name in df.columns:
        df['Translated'] = df[column_name].apply(lambda x: translate_text(x, target_lang))
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            df.to_excel(output_path, index=False)
            messagebox.showinfo("Translation Complete", "Translation and file saved successfully!")
        else:
            messagebox.showerror("Error", "File save operation cancelled.")
    else:
        messagebox.showerror("Column Not Found", f"Column '{column_name}' not found in DataFrame.")

def select_target_file():
    target_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if target_file:
        target_file_path.delete(0, tk.END)
        target_file_path.insert(0, target_file)
        messagebox.showinfo("Selected File", f"Selected file: {target_file}")

# UI setup
root = tk.Tk()
root.title("Excel Translator")
root.geometry("400x300")  # Set initial width and height

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

label = tk.Label(root, text="Enter Column Name:")
label.grid(row=2, column=0,columnspan=2, padx=5, sticky="w")

entry = tk.Entry(root, width=30)
entry.grid(row=2, column=1, padx=5, sticky="ew")

button = tk.Button(root, text="Translate", command=lambda: translate_and_save(target_file_path.get()))
button.grid(row=5, column=0, columnspan=2, pady=10)

var = tk.StringVar(root)
var.set("fr")  # Default target language: French

language_label = tk.Label(root, text="Select Target Language:")
language_label.grid(row=4, column=0, padx=10, sticky="w")

options = [
    'af', 'sq', 'am', 'ar', 'hy', 'az', 'eu', 'be', 'bn', 'bs', 'bg', 'ca', 'ceb', 'ny', 'zh-cn', 'zh-tw', 'co',
    'hr', 'cs', 'da', 'nl', 'en', 'eo', 'et', 'tl', 'fi', 'fr', 'fy', 'gl', 'ka', 'de', 'el', 'gu', 'ht', 'ha', 
    'haw', 'he', 'hi', 'hmn', 'hu', 'is', 'ig', 'id', 'ga', 'it', 'ja', 'jv', 'kn', 'kk', 'km', 'rw', 'ko', 'ku', 
    'ky', 'lo', 'la', 'lv', 'lt', 'lb', 'mk', 'mg', 'ms', 'ml', 'mt', 'mi', 'mr', 'mn', 'my', 'ne', 'no', 'or', 
    'ps', 'fa', 'pl', 'pt', 'pa', 'ro', 'ru', 'sm', 'gd', 'sr', 'st', 'sn', 'sd', 'si', 'sk', 'sl', 'so', 'es', 
    'su', 'sw', 'sv', 'tg', 'ta', 'tt', 'te', 'th', 'tr', 'tk', 'uk', 'ur', 'ug', 'uz', 'vi', 'cy', 'xh', 'yi', 
    'yo', 'zu'
    ]  # Add more languages as needed
dropdown = tk.OptionMenu(root, var, *options)
dropdown.grid(row=4, column=1, sticky="ew")

root.mainloop()
