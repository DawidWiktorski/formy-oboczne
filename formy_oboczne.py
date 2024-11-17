import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from docx import Document
import re
from collections import defaultdict

class DocxAnalyzer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Wyszukiwarka form obocznych")
        self.root.geometry("800x600")
        
        self.docx_path = tk.StringVar()
        self.dict_path = tk.StringVar()
        
        self.create_gui()
        
    def create_gui(self):
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=10, fill='x')
        
        tk.Label(file_frame, text="Wybierz plik DOCX:").pack(side='left')
        tk.Entry(file_frame, textvariable=self.docx_path, width=50).pack(side='left', padx=5)
        tk.Button(file_frame, text="Wybierz", command=self.select_docx).pack(side='left')
        
        dict_frame = tk.Frame(self.root)
        dict_frame.pack(pady=5, padx=10, fill='x')
        tk.Label(dict_frame, text="Wybierz plik słownika:").pack(side='left')
        tk.Entry(dict_frame, textvariable=self.dict_path, width=50).pack(side='left', padx=5)
        tk.Button(dict_frame, text="Wybierz", command=self.select_dict).pack(side='left')
        
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="Szukaj", command=self.analyze_text).pack(side='left', padx=5)
        tk.Button(button_frame, text="Zapisz wyniki", command=self.save_report).pack(side='left', padx=5)
        
        self.result_text = scrolledtext.ScrolledText(self.root, width=80, height=30)
        self.result_text.pack(pady=10, padx=10, fill='both', expand=True)
        
    def select_docx(self):
        filename = filedialog.askopenfilename(filetypes=[("Dokumenty Word", "*.docx")])
        if filename:
            self.docx_path.set(filename)
            
    def select_dict(self):
        filename = filedialog.askopenfilename(filetypes=[("Pliki tekstowe", "*.txt")])
        if filename:
            self.dict_path.set(filename)

    def save_report(self):
        if not self.result_text.get(1.0, tk.END).strip():
            messagebox.showwarning("Uwaga!", "Nie ma danych do zapisania, najpierw przeprowadź analizę.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Plik tekstowy", "*.txt")],
            title="Zapisz wyniki"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(self.result_text.get(1.0, tk.END))
                messagebox.showinfo("Sukces", "Wyniki zostały zapisane.")
            except Exception as e:
                messagebox.showerror("Błąd", f"Wystąpił błąd podczas zapisywania pliku: {str(e)}")

    def get_words_from_dictionary(self, dict_text):
        brackets = re.findall(r'\(([\w,\s]+)\)', dict_text)
        word_groups = []
        
        for bracket in brackets:
            if ',' in bracket and not any(x in bracket.lower() for x in ['l. mn.', 'l. poj.']):
                words = [word.strip() for word in bracket.split(',')]
                if len(words) >= 2:
                    word_groups.append(words)
        
        return word_groups

    def find_words_in_text(self, text, word_group):
        text = text.lower()
        text = re.sub(r'[^\w\s]', ' ', text)
        text_words = set(text.split())
        
        found_words = []
        for word in word_group:
            if word.lower() in text_words:
                found_words.append(word)
                
        return found_words

    def analyze_text(self):
        if not self.docx_path.get() or not self.dict_path.get():
            messagebox.showerror("Błąd", "Musisz wybrać tekst do analizy i słownik!")
            return
            
        try:
            with open(self.dict_path.get(), 'r', encoding='utf-8') as f:
                dict_text = f.read()
            
            word_groups = self.get_words_from_dictionary(dict_text)
            
            doc = Document(self.docx_path.get())
            full_text = ' '.join([paragraph.text for paragraph in doc.paragraphs])
            
            results = []
            for group in word_groups:
                found_words = self.find_words_in_text(full_text, group)
                if len(found_words) >= 2:
                    results.append({
                        'original_group': group,
                        'found_words': found_words
                    })
            
            self.result_text.delete(1.0, tk.END)
            if results:
                self.result_text.insert(tk.END, "Znalezione formy oboczne:\n\n")
                for result in results:
                    self.result_text.insert(tk.END,
                        f"Grupa ze słownika: ({', '.join(result['original_group'])})\n"
                        f"Znaleziono w tekście: {', '.join(result['found_words'])}\n\n")
            else:
                self.result_text.insert(tk.END, 
                    "Nie znaleziono form obocznych, które znajdują się w słowniku.")
                    
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd: {str(e)}")
            
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = DocxAnalyzer()
    app.run()