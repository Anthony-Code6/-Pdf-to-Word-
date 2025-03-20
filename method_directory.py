import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from print_result import log_mensaje
import subprocess
import platform

sistema_operativo = platform.system()

if sistema_operativo == 'Windows':
    import comtypes.client

def seleccionar_carpeta_origen(entrada_origen):
    carpeta_origen = filedialog.askdirectory()
    if carpeta_origen:
        entrada_origen.config(state="normal")
        entrada_origen.delete(0, tk.END)
        entrada_origen.insert(0, carpeta_origen)
        entrada_origen.config(state="disabled")

def seleccionar_carpeta_destino(entrada_destino):
    carpeta_destino = filedialog.askdirectory()
    if carpeta_destino:
        entrada_destino.config(state="normal")
        entrada_destino.delete(0, tk.END)
        entrada_destino.insert(0, carpeta_destino)
        entrada_destino.config(state="disabled")

def convertir_pdf_a_word(entrada_origen, entrada_destino, text_area):
    carpeta_origen = entrada_origen.get()
    carpeta_destino = entrada_destino.get()
    
    if not carpeta_origen or not carpeta_destino:
        messagebox.showerror("Error", "You must select both routes")
        return

    log_mensaje("Starting PDF to Word conversion ...", text_area)

    for archivo in os.listdir(carpeta_origen):
        if archivo.lower().endswith('.pdf'):
            archivo_pdf = os.path.join(carpeta_origen, archivo)
            archivo_docx = os.path.join(carpeta_destino, archivo.replace('.pdf', '.docx'))
            
            log_mensaje(f"Converting: {archivo}...", text_area)
            try:
                cv = Converter(archivo_pdf)
                cv.convert(archivo_docx, start=0, end=None)
                cv.close()
                log_mensaje(f"✅ Converted: {archivo} to Docx.", text_area)
            except Exception as e:
                log_mensaje(f"❌ Error converting {archivo}: {e}", text_area)

    log_mensaje("PDF to Word conversion complete.", text_area)

def convertir_docx_a_pdf(entrada_origen, entrada_destino, text_area):
    carpeta_origen = entrada_origen.get()
    carpeta_destino = entrada_destino.get()

    if not carpeta_origen or not carpeta_destino:
        messagebox.showerror("Error", "You must select both routes")
        return

    log_mensaje("Starting Word to PDF conversion...", text_area)

    try:
        if sistema_operativo == 'Windows':
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False

            for archivo in os.listdir(carpeta_origen):
                if archivo.lower().endswith('.docx'):
                    archivo_docx = os.path.abspath(os.path.join(carpeta_origen, archivo))
                    archivo_pdf = os.path.abspath(os.path.join(carpeta_destino, archivo.replace('.docx', '.pdf')))
                    
                    log_mensaje(f"Converting: {archivo}...", text_area)
                    doc = word.Documents.Open(archivo_docx)
                    doc.SaveAs(archivo_pdf, FileFormat=17)
                    doc.Close()
                    log_mensaje(f"✅ Converted: {archivo} a PDF.", text_area)
            
            word.Quit()

        elif sistema_operativo == 'Linux':
            for archivo in os.listdir(carpeta_origen):
                if archivo.lower().endswith('.docx'):
                    archivo_docx = os.path.abspath(os.path.join(carpeta_origen, archivo))
                    log_mensaje(f"Converting: {archivo}...", text_area)
                    comando = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', carpeta_destino, archivo_docx]
                    subprocess.run(comando, check=True)
                    log_mensaje(f"✅ Converted: {archivo} a PDF.", text_area)
                    
        else:
            messagebox.showerror("Error", f"SO not supported: {sistema_operativo}")

    except Exception as e:
        log_mensaje(f"❌ Conversion error: {e}", text_area)

    log_mensaje("Conversión Word a PDF finalizada.", text_area)
