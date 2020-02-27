import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials

### Google API client authorization ###
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapi.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('coupon-generator.json', scope)
client = gspread.authorize(creds)

sheet = client.open('Coupon Schedule').sheet1

#### GUI ####
win = tk.Tk()
win.title("Coupon HTML Creator")
win.resizable(False, False)

## Notebook
nb = ttk.Notebook(win)
nb.grid(row=2, column=0, pady=(5,0), padx=(1,0), sticky='NSEW')

# Main tab
win = ttk.Frame(nb)
nb.add(win, text="Main")

# Help tab
guide = ttk.Frame(nb)
nb.add(guide, text="Help")

## Frame 1 (Flavor frame)
frame1 = ttk.Frame(win)
frame1.grid(column=0, row=1, sticky="nsew")

flavorLabel = ttk.Label(frame1, text="Flavor Text")
flavorLabel.grid(column=0, row=0, rowspan=2, sticky="w")

flavorDesktopLabel = ttk.Label(frame1, text="Desktop")
flavorDesktopLabel.grid(column=1, row=0, padx=20, sticky="w")

flavorMobileLabel = ttk.Label(frame1, text="Mobile")
flavorMobileLabel.grid(column=1, row=1, padx=20, sticky="w")

flavorDesktopEntryVar = tk.StringVar()
flavorDesktopEntry = ttk.Entry(frame1, width=20, textvariable=flavorDesktopEntryVar)
flavorDesktopEntry.grid(column=2, row=0, padx=(0,20), pady=2, sticky="w")

flavorMobileEntryVar = tk.StringVar()
flavorMobileEntry = ttk.Entry(frame1, width=20, textvariable=flavorMobileEntryVar)
flavorMobileEntry.grid(column=2, row=1, padx=(0,20), pady=2, sticky="w")

## Frame 2 (Coupon 1 frame)
frame2 = ttk.Frame(win)
frame2.grid(column=0, row=2)

coupon1Label = ttk.Label(frame2, text="Coupon 1")
coupon1Label.grid(column=0, row=0, padx=(0,60))

coupon1ImgVar = tk.StringVar()
coupon1Img = ttk.Combobox(frame2, textvariable=coupon1ImgVar)
coupon1Img["values"] = ["image", "from", "google sheets"]
coupon1Img.grid(column=2, row=0, padx=(32,20), pady=2)

# Label for EID (Coupon 1)
c1EIDLabel = ttk.Label(frame2, text="EID")
c1EIDLabel.grid(column=3, row=0, padx=10, sticky="w")

# Label for Timing (Coupon 1)
c1TimeLabel = ttk.Label(frame2, text="Hour (24hr Format)")
c1TimeLabel.grid(column=3, row=1, padx=10, sticky="w")

# Entry for EID (Coupon 1)
c1e1Var = tk.StringVar()
c1e1Entry = ttk.Entry(frame2, width=5, textvariable=c1e1Var)
c1e1Entry.grid(column=4, row=0, padx=(0,2), pady=2, sticky="w")

c1e2Var = tk.StringVar()
c1e2Entry = ttk.Entry(frame2, width=5, textvariable=c1e2Var)
c1e2Entry.grid(column=5, row=0, padx=(0,2), pady=2, sticky="w")

c1e3Var = tk.StringVar()
c1e3Entry = ttk.Entry(frame2, width=5, textvariable=c1e3Var)
c1e3Entry.grid(column=6, row=0, padx=(0,2), pady=2, sticky="w")

# Entry for timing (Coupon 1)
c1t1Var = tk.StringVar()
c1t1Entry = ttk.Entry(frame2, width=5, textvariable=c1t1Var)
c1t1Entry.grid(column=4, row=1, padx=(0,2), pady=2, sticky="w")

c1t2Var = tk.StringVar()
c1t2Entry = ttk.Entry(frame2, width=5, textvariable=c1t2Var)
c1t2Entry.grid(column=5, row=1, padx=(0,2), pady=2, sticky="w")

c1t3Var = tk.StringVar()
c1t3Entry = ttk.Entry(frame2, width=5, textvariable=c1t3Var)
c1t3Entry.grid(column=6, row=1, padx=(0,2), pady=2, sticky="w")

## Frame 3 (Coupon 2 frame)
frame3 = ttk.Frame(win)
frame3.grid(column=0, row=3)

coupon2Label = ttk.Label(frame3, text="Coupon 2")
coupon2Label.grid(column=0, row=0, padx=(0,60))

coupon2ImgVar = tk.StringVar()
coupon2Img = ttk.Combobox(frame3, textvariable=coupon2ImgVar)
coupon2Img["values"] = ["image", "from", "google sheets"]
coupon2Img.grid(column=2, row=0, padx=(32,20), pady=2)

# Label for EID (Coupon 2)
c2EIDLabel = ttk.Label(frame3, text="EID")
c2EIDLabel.grid(column=3, row=0, padx=10, sticky="w")

# Label for Timing (Coupon 2)
c2TimeLabel = ttk.Label(frame3, text="Hour (24hr Format)")
c2TimeLabel.grid(column=3, row=1, padx=10, sticky="w")

# Entry for EID (Coupon 2)
c2e1Var = tk.StringVar()
c2e1Entry = ttk.Entry(frame3, width=5, textvariable=c2e1Var)
c2e1Entry.grid(column=4, row=0, padx=(0,2), pady=2, sticky="w")

c2e2Var = tk.StringVar()
c2e2Entry = ttk.Entry(frame3, width=5, textvariable=c1e2Var)
c2e2Entry.grid(column=5, row=0, padx=(0,2), pady=2, sticky="w")

c2e3Var = tk.StringVar()
c2e3Entry = ttk.Entry(frame3, width=5, textvariable=c2e3Var)
c2e3Entry.grid(column=6, row=0, padx=(0,2), pady=2, sticky="w")

# Entry for timing (Coupon 2)
c2t1Var = tk.StringVar()
c2t1Entry = ttk.Entry(frame3, width=5, textvariable=c2t1Var)
c2t1Entry.grid(column=4, row=1, padx=(0,2), pady=2, sticky="w")

c2t2Var = tk.StringVar()
c2t2Entry = ttk.Entry(frame3, width=5, textvariable=c2t2Var)
c2t2Entry.grid(column=5, row=1, padx=(0,2), pady=2, sticky="w")

c2t3Var = tk.StringVar()
c2t3Entry = ttk.Entry(frame3, width=5, textvariable=c2t3Var)
c2t3Entry.grid(column=6, row=1, padx=(0,2), pady=2, sticky="w")

## Frame 4 (Coupon 3 frame)
frame4 = ttk.Frame(win)
frame4.grid(column=0, row=4)

coupon3Label = ttk.Label(frame4, text="Coupon 3")
coupon3Label.grid(column=0, row=0, padx=(0,60))

coupon3ImgVar = tk.StringVar()
coupon3Img = ttk.Combobox(frame4, textvariable=coupon3ImgVar)
coupon3Img["values"] = ["image", "from", "google sheets"]
coupon3Img.grid(column=2, row=0, padx=(32,20), pady=2)

# Label for EID (Coupon 3)
c3EIDLabel = ttk.Label(frame4, text="EID")
c3EIDLabel.grid(column=3, row=0, padx=10, sticky="w")

# Label for Timing (Coupon 3)
c3TimeLabel = ttk.Label(frame4, text="Hour (24hr Format)")
c3TimeLabel.grid(column=3, row=1, padx=10, sticky="w")

# Entry for EID (Coupon 3)
c3e1Var = tk.StringVar()
c3e1Entry = ttk.Entry(frame4, width=5, textvariable=c3e1Var)
c3e1Entry.grid(column=4, row=0, padx=(0,2), pady=2, sticky="w")

c3e2Var = tk.StringVar()
c3e2Entry = ttk.Entry(frame4, width=5, textvariable=c3e2Var)
c3e2Entry.grid(column=5, row=0, padx=(0,2), pady=2, sticky="w")

c3e3Var = tk.StringVar()
c3e3Entry = ttk.Entry(frame4, width=5, textvariable=c3e3Var)
c3e3Entry.grid(column=6, row=0, padx=(0,2), pady=2, sticky="w")

# Entry for timing (Coupon 3)
c3t1Var = tk.StringVar()
c3t1Entry = ttk.Entry(frame4, width=5, textvariable=c3t1Var)
c3t1Entry.grid(column=4, row=1, padx=(0,2), pady=2, sticky="w")

c3t2Var = tk.StringVar()
c3t2Entry = ttk.Entry(frame4, width=5, textvariable=c3t2Var)
c3t2Entry.grid(column=5, row=1, padx=(0,2), pady=2, sticky="w")

c3t3Var = tk.StringVar()
c3t3Entry = ttk.Entry(frame4, width=5, textvariable=c3t3Var)
c3t3Entry.grid(column=6, row=1, padx=(0,2), pady=2, sticky="w")

## Labels
helpMsg = "WIP"
helpLabel = tk.Message(guide, text=helpMsg, width=500)
helpLabel.grid(row=0, column=0)

imgLabel = ttk.Label(win, text="Image URL")
imgLabel.grid(column=0, row=0, padx=180, sticky="w")






win.mainloop()