import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from docx import Document

from doc_utils import add_toc, update_toc
from extractor import extractInfo


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    selected_path = filedialog.askdirectory(title="Selecciona la carpeta del proyecto")
    root.destroy()

    if not selected_path:
        print("No se seleccionó ninguna carpeta.")
        exit()

    folder_name = os.path.basename(selected_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = f"codigo_{folder_name}_{timestamp}.docx"
    document = Document()
    document.add_heading("Tabla de contenidos", level=0)
    add_toc(document)
    document.add_page_break()
    extractInfo(selected_path, document)
    document.save(output)
    update_toc(output)
