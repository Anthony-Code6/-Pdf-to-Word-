import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os
import comtypes.client  # Para convertir DOCX a PDF con Microsoft Word

# Obtener la ruta de la carpeta Descargas del usuario
ruta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")

# Función para seleccionar carpeta de origen
def seleccionar_carpeta_origen():
    carpeta_origen = filedialog.askdirectory()
    if carpeta_origen:
        entrada_origen.config(state="normal")
        entrada_origen.delete(0, tk.END)
        entrada_origen.insert(0, carpeta_origen)
        entrada_origen.config(state="disabled")

# Función para seleccionar carpeta de destino
def seleccionar_carpeta_destino():
    carpeta_destino = filedialog.askdirectory()
    if carpeta_destino:
        entrada_destino.config(state="normal")
        entrada_destino.delete(0, tk.END)
        entrada_destino.insert(0, carpeta_destino)
        entrada_destino.config(state="disabled")

# Función para convertir PDF a DOCX
def convertir_pdf_a_word():
    carpeta_origen = entrada_origen.get()
    carpeta_destino = entrada_destino.get()
    
    if not carpeta_origen or not carpeta_destino:
        messagebox.showerror("Error", "Deben seleccionar las rutas de origen y destino")
        return
    
    for archivo in os.listdir(carpeta_origen):
        if archivo.lower().endswith('.pdf'):
            archivo_pdf = os.path.join(carpeta_origen, archivo)
            archivo_destino = os.path.join(carpeta_destino, archivo.replace('.pdf', '.docx'))
            
            # Convertir PDF a DOCX
            cv = Converter(archivo_pdf)
            cv.convert(archivo_destino, start=0, end=None)
            cv.close()
            print(f"✅ Convertido: {archivo_pdf} -> {archivo_destino}")

# Función mejorada para convertir DOCX a PDF (soporta imágenes y nombres con espacios)
def convertir_docx_a_pdf():
    carpeta_origen = entrada_origen.get()
    carpeta_destino = entrada_destino.get()

    if not carpeta_origen or not carpeta_destino:
        messagebox.showerror("Error", "Deben seleccionar las rutas de origen y destino")
        return

    try:
        # Intentar abrir Microsoft Word
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False  # No mostrar Word

        for archivo in os.listdir(carpeta_origen):
            if archivo.lower().endswith('.docx'):
                archivo_docx = os.path.join(carpeta_origen, archivo)
                
                # Crear la ruta de destino con el mismo nombre pero en formato PDF
                archivo_pdf = os.path.join(carpeta_destino, archivo.replace('.docx', '.pdf'))
                
                try:
                    # Reemplazar las barras invertidas para evitar problemas en rutas largas o con espacios
                    archivo_docx = os.path.abspath(archivo_docx)
                    archivo_pdf = os.path.abspath(archivo_pdf)

                    print(f"📄 Procesando: {archivo_docx} -> {archivo_pdf}")

                    # Abrir el documento en Word
                    doc = word.Documents.Open(archivo_docx)
                    
                    # Guardar el documento como PDF
                    doc.SaveAs(archivo_pdf, FileFormat=17)  
                    doc.Close()  # Cerrar el documento después de la conversión
                    
                    print(f"✅ Convertido con éxito: {archivo_docx} -> {archivo_pdf}")

                except Exception as e:
                    print(f"❌ Error al convertir {archivo_docx}:\n{str(e)}")
                    messagebox.showerror("Error", f"No se pudo convertir {archivo_docx} a PDF.\n\n{str(e)}")

        # Cerrar Word completamente
        word.Quit()
        
    except Exception as e:
        messagebox.showerror("Error", "Microsoft Word no está instalado o hubo un problema al iniciar Word.")
        print(f"❌ Error al iniciar Word: {str(e)}")

# Configuración de la interfaz gráfica con Tkinter
ventana = tk.Tk()
ventana.title("Convertir PDF a Word / DOCX a PDF")
# ventana.iconbitmap("icono.ico")
ventana.resizable(False, False)
ventana.geometry("420x150")

# Crear un frame para organizar las entradas y botones
frame_inputs = tk.Frame(ventana)
frame_inputs.pack(pady=10)

# Primera fila: Carpeta de origen
tk.Label(frame_inputs, text="Carpeta de origen:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entrada_origen = tk.Entry(frame_inputs, width=40, state="disabled")
entrada_origen.grid(row=0, column=1, padx=5, pady=5)
tk.Button(frame_inputs, text="+", command=seleccionar_carpeta_origen).grid(row=0, column=2, padx=5, pady=5)

# Segunda fila: Carpeta de destino
tk.Label(frame_inputs, text="Carpeta de destino:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entrada_destino = tk.Entry(frame_inputs, width=40)
entrada_destino.grid(row=1, column=1, padx=5, pady=5)
entrada_destino.insert(0, ruta_descargas)  # Inicializa con la carpeta Descargas
tk.Button(frame_inputs, text="+", command=seleccionar_carpeta_destino).grid(row=1, column=2, padx=5, pady=5)

# Frame contenedor para los botones de conversión
frame_botones = tk.Frame(ventana)
frame_botones.pack(pady=10)

# Botón para convertir PDF a Word
tk.Button(frame_botones, text="Convertir PDF a Word", command=convertir_pdf_a_word).pack(side="left", pady=5, padx=10)

# Botón para convertir DOCX a PDF
tk.Button(frame_botones, text="Convertir Word a PDF", command=convertir_docx_a_pdf).pack(side="left", pady=5, padx=10)

# Iniciar la ventana
ventana.mainloop()

