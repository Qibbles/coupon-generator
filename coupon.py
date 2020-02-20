import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

#### GUI ####
win = tk.Tk()
win.title("Coupon HTML Creator")
win.resizable(False, False)

# Notebook
nb = ttk.Notebook(win)
nb.grid(row=2, column=0, pady=(5,0))