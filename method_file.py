import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os
import platform
import comtypes.client
import subprocess

from print_result import log_mensaje

sistema_operativo = platform.system()
ruta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")

def seleccionar_archivo(entrada_file):
    file = filedialog.askopenfilename(
        title="Seleccionar archivo",
        filetypes=[
            ("Documentos", "*.pdf *.docx *.doc *.docm *.odt"),
            ("PDF files", "*.pdf"),
            ("Word Documents", "*.docx *.doc *.docm"),
            ("OpenDocument Text", "*.odt"),
            ("Todos los archivos", "*.*")
        ]
    )
    if file:
        entrada_file.config(state="normal")
        entrada_file.delete(0, tk.END)
        entrada_file.insert(0, file)
        entrada_file.config(state="disabled")

def convertir_pdf_to_word(entrada_file, text_area):
    archivo_origen = entrada_file.get()
    
    if not archivo_origen.lower().endswith('.pdf'):
        messagebox.showerror("Error", "The selected file is not a PDF")
        return

    log_mensaje("Starting PDF to Word conversion ...", text_area)

    archivo_nombre = os.path.basename(archivo_origen)
    archivo_docx = os.path.join(ruta_descargas, archivo_nombre.replace('.pdf', '.docx'))
    
    log_mensaje(f"Converting: {archivo_nombre}...", text_area)
    try:
        cv = Converter(archivo_origen)
        cv.convert(archivo_docx, start=0, end=None)
        cv.close()
        log_mensaje(f"✅ Converted: {archivo_nombre} to Docx.", text_area)
    except Exception as e:
        log_mensaje(f"❌ Error converting {archivo_nombre}: {e}", text_area)

    log_mensaje("PDF to Word conversion complete.", text_area)

def convertir_docx_to_pdf(entrada_file, text_area):
    archivo_origen = entrada_file.get()
    
    if not archivo_origen.lower().endswith('.docx'):
        messagebox.showerror("Error", "The selected file is not a DOCX")
        return

    log_mensaje("Starting Word to PDF conversion...", text_area)

    archivo_nombre = os.path.basename(archivo_origen)
    archivo_pdf = os.path.abspath(os.path.join(ruta_descargas, archivo_nombre.replace('.docx', '.pdf')))
    
    try:
        if sistema_operativo == 'Windows':
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False

            log_mensaje(f"Converting: {archivo_nombre}...", text_area)
            doc = word.Documents.Open(os.path.abspath(archivo_origen))
            doc.SaveAs(archivo_pdf, FileFormat=17)
            doc.Close()
            word.Quit()
            log_mensaje(f"✅ Converted: {archivo_nombre} a PDF.", text_area)
        
        elif sistema_operativo == 'Linux':
            log_mensaje(f"Converting: {archivo_nombre}...", text_area)
            comando = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', ruta_descargas, os.path.abspath(archivo_origen)]
            subprocess.run(comando, check=True)
            log_mensaje(f"✅ Converted: {archivo_nombre} a PDF.", text_area)
        
        else:
            messagebox.showerror("Error", f"SO not supported: {sistema_operativo}")

    except Exception as e:
        log_mensaje(f"❌ Conversion error: {e}", text_area)

    log_mensaje("Conversión Word a PDF finalizada.", text_area)