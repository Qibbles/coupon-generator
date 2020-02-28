import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials

### Google API client authorization ###
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('coupon-generator.json', scope)
client = gspread.authorize(creds)

sheet = client.open('Coupon Schedule')
worksheet = sheet.worksheet("Coupon Summary")

### Functions ###
def generate():
    coupon1 = coupon1ImgVar.get()
    coupon2 = coupon2ImgVar.get()
    coupon3 = coupon3ImgVar.get()

    if len(coupon1) == 0:
        messagebox.showinfo("Oops!","At least 1 Coupon image is required!")
    else:
        flavorDesktop = flavorDesktopEntryVar.get()
        flavorMobile = flavorMobileEntryVar.get()
        # Coupon 1 variables
        c1e1 = c1e1Var.get()
        c1e2 = c1e2Var.get()
        c1e3 = c1e3Var.get()
        c1t1 = c1t1Var.get()
        c1t2 = c1t2Var.get()
        c1t3 = c1t3Var.get()
        # Coupon 2 variables
        c2e1 = c2e1Var.get()
        c2e2 = c2e2Var.get()
        c2e3 = c2e3Var.get()
        c2t1 = c2t1Var.get()
        c2t2 = c2t2Var.get()
        c2t3 = c2t3Var.get()
        # Coupon 3 variables
        c3e1 = c3e1Var.get()
        c3e2 = c3e2Var.get()
        c3e3 = c3e3Var.get()
        c3t1 = c3t1Var.get()
        c3t2 = c3t2Var.get()
        c3t3 = c3t3Var.get()

        html = open("coupon.html", "w")
        html.write('<style>\n')
        html.write('.button {background-color: #f8f8fa; border: none;color: white; padding: 12px 35px; text-align: center;text-decoration: none; display: inline-block; font-size: 14px; margin: 4px 2px; -webkit-transition-duration: 0.4s; /* Safari */  transition-duration: 0.4s; cursor: pointer;}\n')
        html.write('.buttonCpn {background-color: #ffffff; color: black; border: 1px solid #dbe1e2;border-radius: 0px;}\n')
        html.write('.buttonCpn:hover {background-color: #f8f8fa; color:black;}\n')
        html.write('</style>\n')
        html.write('\n')
        html.write('<script>\n')
        html.write('function eventApplyTime(x) {\n')
        html.write('    var currentHour = (new Date()).getHours();\n')
        html.write('    switch(x) {\n')
        html.write('        case 1:\n')
        html.write('            var events = new Array("' + c1e1 + '","' + c1e2 + '","' + c1e3 + '")\n')
        html.write('            var hours = new Array("' + c1t1 + '","' + c1t2 +'","' + c1t3 + '");\n')
        html.write('            break;\n')
        html.write('        case 2:\n')
        html.write('            var events = new Array("' + c2e1 + '","' + c2e2 + '","' + c2e3 + '")\n')
        html.write('            var hours = new Array("' + c2t1 + '","' + c2t2 +'","' + c2t3 + '");\n')
        html.write('            break;\n')
        html.write('        case 3:\n')
        html.write('            var events = new Array("' + c3e1 + '","' + c3e2 + '","' + c3e3 + '")\n')
        html.write('            var hours = new Array("' + c3t1 + '","' + c3t2 +'","' + c3t3 + '");\n')
        html.write('            break;\n')
        html.write('    }\n')
        html.write('    if (hours.length = 1) {Util.EventApply(events[0]);}\n')
        html.write('        else if (hours.length = 2) {\n')
        html.write('            else if (currentHour >=hours[1]) {Util.EventApply(events[1]);}\n')
        html.write('        }\n')
        html.write('        else if (hours.length = 3) {\n')
        html.write('            if (currentHour >= hours[0] && currentHour < hours[1]) {Util.EventApply(events[0]);}\n')
        html.write('                else if (currentHour >=hours[1] && currentHour < hours[2]) {Util.EventApply(events[1]);}\n')
        html.write('                else if (currentHour >=hours[2] ) {Util.EventApply(events[2]);}\n')
        html.write('    }\n')
        html.write('};\n')
        html.write('</script>\n')
        html.write('\n')
        html.write('<table width="100%" border="0" cellpadding="0" cellspacing="0">\n')
        html.write('    <tr>\n')
        html.write('        <td><img src="'+ flavorDesktop + '" width="100%" alt=""></a></td>\n')
        html.write('        <td><a href="javascript:eventApplyTime(1)"><img src="' + coupon1 + '" width="100%" alt=""></a></td>\n')
        html.write('        <td><a href="javascript:eventApplyTime(2)"><img src="' + coupon2 + '" width="100%" alt=""></a></td>\n')
        html.write('    </tr>\n')
        html.write('</table>\n')
        html.write('\n')
        html.write('<div align="center" style="padding: 15px 5px;">\n')
        html.write('<a href="https://dp.image-gmkt.com/dp2016/SG/design/campaign/2020/03_Mar/0301/coupons/0301_2Coupon_TnCs.jpg" onClick="window.open("https://dp.image-gmkt.com/dp2016/SG/design/campaign/2020/03_Mar/0301/coupons/0301_2Coupon_TnCs.jpg","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=no, width=600,height=610,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">Terms and Conditions &#9656;</button></a>\n')
        html.write('<a href="https://www.qoo10.sg/gmkt.inc/Event/qchance.aspx" target="_blank"><button class="button buttonCpn">Get MameQ and Rewards &#9656;</button></a>\n')
        html.write('<a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/07/sendcoupon_WEB.html?2" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/07/sendcoupon_WEB.html?2","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=1, width=950,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">Send Coupon &#9656;</button></a>\n')
        html.write('<a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/11/howtousecoupon_WEB.html?2" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/11/howtousecoupon_WEB.html?2","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=1, width=950,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">How to use Coupon? &#9656;</button></a>\n')
        html.write('</div>\n')
        html.write('<br>\n')
        html.close()
    # if len(coupon2) == 0:
    #     print("No coupon 2")
    # if len(coupon3) == 0:
    #     print("No coupon 3")
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

# coupon1Img["values"] = worksheet.cell(23, 1).value
print(worksheet.get_all_values()[0])
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
c2e2Entry = ttk.Entry(frame3, width=5, textvariable=c2e2Var)
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

## Frame 5 (Buttons frame)
frame5 = ttk.Frame(win)
frame5.grid(column=0, row=5)

# Button
generateButton = ttk.Button(frame5, text="Generate", command=generate)
generateButton.grid(column=0, row=0)

## Labels
helpMsg = "WIP"
helpLabel = tk.Message(guide, text=helpMsg, width=500)
helpLabel.grid(row=0, column=0)

imgLabel = ttk.Label(win, text="Image URL")
imgLabel.grid(column=0, row=0, padx=180, sticky="w")







win.mainloop()