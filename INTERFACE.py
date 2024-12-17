import tkinter as tk
from tkinter import *
from tkinter import ttk


def calculate(*args):
    try:
        value = float(feet.get())
        meters.set(int(0.3048 * value * 10000.0 + 0.5)/10000.0)
    except ValueError:
        pass

root = Tk()
root.title("Programa de descarga")

mainframe = ttk.Frame(root, padding = "3 3 12 12")
mainframe.config(bg = "lightblue")
mainframe.grid(column = 0, row = 0, sticky = (N, W, E, S))
root.columnconfigure(0, weight = 1)
root.rowconfigure(0, weight = 1)

usuario = StringVar()
usuario_entry = ttk.Entry(mainframe, width = 7, textvariable = usuario)
usuario_entry.grid(column = 2, row = 1, sticky = (W, E))
contrasenna = StringVar()
contrasenna_entry = ttk.Entry(mainframe, width = 7, textvariable = contrasenna)
contrasenna_entry.grid(column = 2, row = 2, sticky = (W, E))

ttk.Button(mainframe, text = "Iniciar", command = calculate).grid(column = 3, row = 3, sticky = W)

ttk.Label(mainframe, text = "Usuario").grid(column = 1, row = 1, sticky = W)
ttk.Label(mainframe, text = "Contraseña").grid(column = 1, row=2, sticky=W)

for child in mainframe.winfo_children(): 
    child.grid_configure(padx=5, pady=5)

usuario_entry.focus()
root.bind("<Return>", calculate)

root.mainloop()

