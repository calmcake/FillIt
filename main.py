import tkinter as tk
from tkinter import filedialog, messagebox
from docxtpl import DocxTemplate
import docx
import re
import os

entries = {}
template_path = ""
fields_frame = None

def load_template():
    global template_path, entries, fields_frame

    filepath = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if not filepath:
        return

    template_path = filepath
    entries.clear()

    # Очистка старых полей
    for widget in fields_frame.winfo_children():
        widget.destroy()

    doc = docx.Document(filepath)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    matches = [re.findall(r'\{\{(.*?)\}\}', line) for line in paragraphs]
    flat_matches = list(set([item.strip() for sublist in matches for item in sublist]))

    for key in flat_matches:
        label = tk.Label(fields_frame, text=f"{key}:")
        label.pack(anchor="w", padx=10)
        entry = tk.Entry(fields_frame, width=50)
        entry.pack(padx=10, pady=5)
        entries[key] = entry

def generate_document():
    if not template_path:
        messagebox.showerror("Ошибка", "Сначала загрузите шаблон.")
        return

    context = {}
    for key, entry in entries.items():
        value = entry.get()
        context[key] = value

    try:
        doc = DocxTemplate(template_path)
        doc.render(context)
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if save_path:
            doc.save(save_path)
            messagebox.showinfo("Успех", f"Документ сохранён: {save_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


# Создание окна
root = tk.Tk()
root.title("Генератор DOCX из шаблона")
root.geometry("500x600")

btn_load = tk.Button(root, text="Загрузить шаблон DOCX", command=load_template)
btn_load.pack(pady=10)

canvas = tk.Canvas(root)
scrollbar = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

fields_frame = scrollable_frame

btn_generate = tk.Button(root, text="Сгенерировать", command=generate_document)
btn_generate.pack(pady=10)

root.mainloop()
