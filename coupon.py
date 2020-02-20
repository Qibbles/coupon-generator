import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

#### GUI ####
win = tk.Tk()
win.title("Coupon HTML Creator")
win.resizable(False, False)

## Notebook
nb = ttk.Notebook(win)
nb.grid(row=2, column=0, pady=(5,0), padx=(1,0), sticky='NSEW')

# Input tab
win = ttk.Frame(nb)
nb.add(win, text="Main")

# Help tab
win = ttk.Frame(nb)
nb.add(win, text="Help")

helpMsg = "WIP"
helpLabel = tk.Message(guide, text=helpMsg, width=500)
helpLabel.grid(row=0, column=0)

