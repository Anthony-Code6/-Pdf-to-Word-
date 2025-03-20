import tkinter as tk
from tkinter import ttk,PhotoImage
import os
import platform
from method_directory import seleccionar_carpeta_destino,seleccionar_carpeta_origen,convertir_docx_a_pdf,convertir_pdf_a_word
from method_file import seleccionar_archivo,convertir_docx_to_pdf,convertir_pdf_to_word

ruta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
sistema_operativo = platform.system()

ventana = tk.Tk()
ventana.title("LegionSoft - Converter (PDF/Word & Word/PDF)")
ventana.resizable(False,False)

if sistema_operativo == 'Windows':
    ventana.geometry("586x380")
    ventana.iconbitmap('icono.ico')
elif sistema_operativo == 'Linux':
    ventana.geometry("789x380")
    imagen=PhotoImage(file='icono.png')
    ventana.tk.call('wm','iconphoto',ventana._w,imagen)

notebook = ttk.Notebook(ventana)
notebook.pack(expand=True, fill='both')

tab_directory = tk.Frame(notebook)
notebook.add(tab_directory, text="Directory")

frame_directory = tk.Frame(tab_directory)
frame_directory.pack(pady=10, padx=10, fill='x')

tk.Label(frame_directory, text="Source Folder:").grid(row=0, column=0, sticky='w')
origen_frame = tk.Frame(frame_directory)
origen_frame.grid(row=1, column=0, sticky='ew', pady=(0, 10))
entrada_origen = tk.Entry(origen_frame, width=90, state="disabled")
entrada_origen.pack(side=tk.LEFT, fill='x', expand=True)
tk.Button(origen_frame, text="+", command=lambda:seleccionar_carpeta_origen(entrada_origen)).pack(side=tk.LEFT)

tk.Label(frame_directory, text="Destination Folder:").grid(row=2, column=0, sticky='w')
destino_frame = tk.Frame(frame_directory)
destino_frame.grid(row=3, column=0, sticky='ew')
entrada_destino = tk.Entry(destino_frame, width=90, state="disabled")
entrada_destino.pack(side=tk.LEFT, fill='x', expand=True)
entrada_destino.insert(0, ruta_descargas)
tk.Button(destino_frame, text="+", command=lambda:seleccionar_carpeta_destino(entrada_destino)).pack(side=tk.LEFT)

frame_botones = tk.Frame(tab_directory)
frame_botones.pack(pady=10)
tk.Button(frame_botones, text="PDF to Word", command=lambda:convertir_pdf_a_word(entrada_origen,entrada_destino,text_area)).pack(side=tk.LEFT, padx=10)
tk.Button(frame_botones, text="Word to PDF", command=lambda:convertir_docx_a_pdf(entrada_origen,entrada_destino,text_area)).pack(side=tk.LEFT, padx=10)



tab_file = tk.Frame(notebook)
notebook.add(tab_file, text="File")

frame_file = tk.Frame(tab_file)
frame_file.pack(pady=10, padx=10, fill='x')

tk.Label(frame_file, text="Choose the document:").grid(row=0, column=0, sticky='w')
origen_frame = tk.Frame(frame_file)
origen_frame.grid(row=1, column=0, sticky='ew', pady=(0, 10))
entrada_file = tk.Entry(origen_frame, width=90, state="disabled")
entrada_file.pack(side=tk.LEFT, fill='x', expand=True)
tk.Button(origen_frame, text="+", command=lambda: seleccionar_archivo(entrada_file)).pack(side=tk.LEFT)

# Frame File

frame_botones_file = tk.Frame(tab_file)
frame_botones_file.pack(pady=10)
tk.Button(frame_botones_file, text="PDF to Word", command=lambda: convertir_pdf_to_word(entrada_file, text_area)).pack(side=tk.LEFT, padx=10)
tk.Button(frame_botones_file, text="Word to PDF", command=lambda: convertir_docx_to_pdf(entrada_file, text_area)).pack(side=tk.LEFT, padx=10)


tk.Label(tab_file, text='They will be downloaded automatically in the following path: '+ruta_descargas).pack(padx=10, pady=10)

# Área de texto para logs
text_area = tk.Text(ventana, height=10, state="disabled")
text_area.pack(fill='both', padx=10, pady=10)

ventana.mainloop()
