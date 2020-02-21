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

coupon1Label = ttk.Label(win, text="Coupon 1")
coupon1Label.grid(column=0, row=0)

coupon2Label = ttk.Label(win, text="Coupon 2")
coupon2Label.grid(column=0, row=1)

coupon3Label = ttk.Label(win, text="Coupon 3")
coupon3Label.grid(column=0, row=2)