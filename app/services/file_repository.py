from tkinter import messagebox
from pdf2docx import Converter
import os
import platform
import subprocess

from app.tools import path_download 
from app.tools import print_message as log

sistema_operativo = platform.system()
ruta_descargas = path_download.obtener_ruta_descargas()

if sistema_operativo == 'Windows':
    import comtypes.client
        

def convertir_pdf_to_word(entrada_file, text_area):
    archivo_origen = entrada_file.get()
    
    if not archivo_origen.lower().endswith('.pdf'):
        messagebox.showerror("Error", "The selected file is not a PDF")
        return

    log.log_mensaje("Starting PDF to Word conversion ...", text_area)

    archivo_nombre = os.path.basename(archivo_origen)
    archivo_docx = os.path.join(ruta_descargas, archivo_nombre.replace('.pdf', '.docx'))
    
    log.log_mensaje(f"Converting: {archivo_nombre}...", text_area)
    try:
        cv = Converter(archivo_origen)
        cv.convert(archivo_docx, start=0, end=None)
        cv.close()
        log.log_mensaje(f"✅ Converted: {archivo_nombre} to Docx.", text_area)
    except Exception as e:
        log.log_mensaje(f"❌ Error converting {archivo_nombre}: {e}", text_area)

    log.log_mensaje("PDF to Word conversion complete.", text_area)

def convertir_docx_to_pdf(entrada_file, text_area):
    archivo_origen = entrada_file.get()
    
    valiate_ext = ('.docx', '.doc', '.docm', '.odt')

    if not archivo_origen.lower().endswith(valiate_ext):
        messagebox.showerror("Error", "The selected file is not a DOCX")
        return

    log.log_mensaje("Starting Word to PDF conversion...", text_area)

    # Obtener el nombre base del archivo y la extensión
    archivo_nombre, extension = os.path.splitext(os.path.basename(archivo_origen))
    
    # Crear el nombre del archivo PDF con la extensión correcta
    archivo_pdf = os.path.abspath(os.path.join(ruta_descargas, archivo_nombre + '.pdf'))

    try:
        if sistema_operativo == 'Windows':
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False

            log.log_mensaje(f"Converting: {archivo_nombre}...", text_area)
            doc = word.Documents.Open(os.path.abspath(archivo_origen))
            doc.SaveAs(archivo_pdf, FileFormat=17)
            doc.Close()
            word.Quit()
            log.log_mensaje(f"✅ Converted: {archivo_nombre} a PDF.", text_area)
        
        elif sistema_operativo == 'Linux':
            log.log_mensaje(f"Converting: {archivo_nombre}...", text_area)
            comando = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', ruta_descargas, os.path.abspath(archivo_origen)]
            subprocess.run(comando, check=True)
            log.log_mensaje(f"✅ Converted: {archivo_nombre} a PDF.", text_area)
        
        else:
            messagebox.showerror("Error", f"SO not supported: {sistema_operativo}")

    except Exception as e:
        log.log_mensaje(f"❌ Conversion error: {e}", text_area)

    log.log_mensaje("Conversión Word a PDF finalizada.", text_area)