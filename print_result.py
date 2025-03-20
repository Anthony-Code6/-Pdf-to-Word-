import tkinter as tk

def log_mensaje(mensaje, text_area):
    text_area.config(state="normal")
    text_area.insert(tk.END, mensaje + "\n")
    text_area.yview_moveto(1.0)  # Mover siempre al final
    text_area.config(state="disabled")
    text_area.update()