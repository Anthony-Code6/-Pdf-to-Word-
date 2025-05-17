import tkinter as tk
from tkinter import filedialog

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