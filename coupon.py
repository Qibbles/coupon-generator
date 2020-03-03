import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials

######## Google API client authorization ########
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('coupon-generator.json', scope)
client = gspread.authorize(creds)

sheet = client.open('Coupon Schedule')
worksheet = sheet.worksheet("Coupon Summary")

######## Functions ########

## Populating design dropdown
def designs():
    global valueDict
    global valueDictKey
    valueDict = {}
    valueDictKey = []
    designList = []

    for i in range(len(worksheet.col_values(1))):
        designList.append(worksheet.cell(i+1, 3).value)
    del designList[0:2]
# combine 2 for loop?
    for i in range(len(designList)):
        if designList[i] != "":
            valueDict[worksheet.cell(i+3, 1).value] = designList[i]
            
    for each in valueDict:
        valueDictKey.append(each)

    print(valueDict)
    print(valueDictKey)

designs()

## Generating HTML
def generate():
    global coupon1EID
    global coupon2EID
    global coupon3EID
    global coupon1Dict
    global coupon2Dict
    global coupon3Dict
    global flavorText
    global date
    noCoupon = []
    coupon1Dict = {}
    coupon2Dict = {}
    coupon3Dict = {}

    if len(coupon1ImgVar.get()) == 0:
        messagebox.showinfo("Oops!","At least 1 Coupon image is required!")
    else:
        if len(coupon1ImgVar.get()) != 0: 
            noCoupon.append(coupon1ImgVar.get())
        if len(coupon2ImgVar.get()) != 0:
            noCoupon.append(coupon2ImgVar.get())   
        if len(coupon3ImgVar.get()) != 0:
            noCoupon.append(coupon3ImgVar.get())

    if len(noCoupon) >= 1:     
            flavorText = flavorTextEntryVar.get()
            date = dateEntryVar.get()
            if coupon1ImgVar.get() not in valueDict:
                # coupon1.append(valueDict[coupon1ImgVar.get()])
            # else:
                # Write coupon design into google sheet
                # coupon1 = coupon1ImgVar.get()
                worksheet.append_row(['Temp','', coupon1ImgVar.get()])
                valueDict[coupon1ImgVar.get()] = coupon1ImgVar.get()
            # Coupon 1 variables
            if len(c1e1Var.get()) == 0:
                messagebox.showinfo("Error", "You require at least 1 EID!")
            else:
                if len(c1e1Var.get()) >= 1:
                    coupon1Dict[c1e1Var.get()] = c1t1Var.get()
                    coupon1EID = 1
                if len(c1e2Var.get()) >= 1:
                    coupon1Dict[c1e2Var.get()] = c1t2Var.get()
                    coupon1EID = 2
                if len(c1e3Var.get()) >= 1:
                    coupon1Dict[c1e3Var.get()] = c1t3Var.get()
                    coupon1EID = 3 
    if len(noCoupon) >= 2:
        if coupon2ImgVar.get() not in valueDict:
            worksheet.append_row(['Temp','', coupon2ImgVar.get()])
            valueDict[coupon2ImgVar.get()] = coupon2ImgVar.get()
        # Coupon 1 variables
        if len(c2e1Var.get()) == 0:
            messagebox.showinfo("Error", "You require at least 1 EID!")
        else:
            if len(c2e1Var.get()) >= 1:
                coupon2Dict[c2e1Var.get()] = c2t1Var.get()
                coupon2EID = 1
            if len(c2e2Var.get()) >= 1:
                coupon2Dict[c2e2Var.get()] = c2t2Var.get()
                coupon2EID = 2
            if len(c2e3Var.get()) >= 1:
                coupon2Dict[c2e3Var.get()] = c2t3Var.get()
                coupon2EID = 3 
    if len(noCoupon) >= 3:
        if coupon3ImgVar.get() not in valueDict:
            worksheet.append_row(['Temp','', coupon3ImgVar.get()])
            valueDict[coupon3ImgVar.get()] = coupon3ImgVar.get()
        # Coupon 1 variables
        if len(c3e1Var.get()) == 0:
            messagebox.showinfo("Error", "You require at least 1 EID!")
        else:
            if len(c3e1Var.get()) >= 1:
                coupon3Dict[c3e1Var.get()] = c3t1Var.get()
                coupon3EID = 1
            if len(c3e2Var.get()) >= 1:
                coupon3Dict[c3e2Var.get()] = c3t2Var.get()
                coupon3EID = 2
            if len(c3e3Var.get()) >= 1:
                coupon3Dict[c3e3Var.get()] = c3t3Var.get()
                coupon3EID = 3 

    HTML(noCoupon)

def HTML(x):
    html = open("Coupon_" + date + ".html", "w")
    ##### Desktop #####
    # CSS 
    html.write('<style>\n')
    html.write('@import url("https://fonts.googleapis.com/css?family=Montserrat:400,500,700&display=swap");\n')
    html.write('.button {background-color: #f8f8fa; border: none;color: white; padding: 12px 35px; text-align: center; text-decoration: none; display: inline-block; font-size: 14px; margin: 4px 2px; -webkit-transition-duration: 0.4s; /* Safari */  transition-duration: 0.4s; cursor: pointer;}\n')
    html.write('.buttonCpn {background-color: #ffffff; color: black; border: 1px solid #dbe1e2;border-radius: 0px;}\n')
    html.write('.buttonCpn:hover {background-color: #f8f8fa; color:black;}\n')
    html.write('\n')
    html.write('.cartcoupontable {\n')
    html.write('border: 1px solid #dbe1e2;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.cartcoupondate {\n')
    html.write('font-family: Montserrat, Helvetica, Arial, sans-serif;\n')
    html.write('background: #ec2e3c;\n')
    html.write('border-radius: 18px;\n')
    html.write('border: 0px;\n')
    html.write('color: #FFF;\n')
    html.write('font-size: 16px;\n')
    html.write('letter-spacing: 1px;\n')
    html.write('padding: 7px 25px;\n')
    html.write('}\n')
    html.write('.cartcouponphrase {\n')
    html.write('font-family: Montserrat, Helvetica, Arial, sans-serif;\n')
    html.write('font-size: 38px;\n')
    html.write('font-weight: 700;\n')
    html.write('color: #333333;\n')
    html.write('line-height: 40px;\n')
    html.write('}\n')
    html.write('</style>\n')
    html.write('\n')
    html.write('<script>\n')
    html.write('function eventApplyTime(x) {\n')
    html.write('    var currentHour = (new Date()).getHours();\n')
    html.write('    switch(x) {\n')
    if len(x) >= 1:
        # Coupon 1 conditionals
        html.write('        case 1:\n')
        eidList = []
        timeList = []
        for eid, time in coupon1Dict.items():
            eidList.append(eid)
            timeList.append(time)
        if coupon1EID == 3:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '","' + str(eidList[2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '","' + str(timeList[2]) + '");\n')
        elif coupon1EID == 2:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '")\n')
        html.write('            break;\n')
        eidList.clear()
        timeList.clear()
    if len(x) >= 2:
        # Coupon 2 conditionals
        html.write('        case 2:\n')
        eid2List = []
        time2List = []
        for eid, time in coupon2Dict.items():
            eidList.append(eid)
            timeList.append(time)
        if coupon2EID == 3:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '","' + str(eidList[2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '","' + str(timeList[2]) + '");\n')
        elif coupon2EID == 2:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '")\n')
        html.write('            break;\n')
        eidList.clear()
        timeList.clear()
    if len(x) >= 3:
        # Coupon 3 conditionals
        html.write('        case 3:\n')
        eid3List = []
        time3List = []
        for eid, time in coupon3Dict.items():
            eidList.append(eid)
            timeList.append(time)
        print(eidList)
        print(timeList)
        if coupon3EID == 3:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '","' + str(eidList[2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '","' + str(timeList[2]) + '");\n')
        elif coupon3EID == 2:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '")\n')
        html.write('            break;\n')
        eidList.clear()
        timeList.clear()
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
    html.write('<table width="980" height="200" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#f8f8fa" class="cartcoupontable">\n')
    html.write('<tr>\n')
    if len(flavorText) and len(date) != 0:
        html.write('    <td width="328" style="padding-left: 26px;">\n')
        html.write('        <span class="cartcoupondate">' + date + '</span><br><br>\n') ########## Flavor date goes here ##########
        html.write('        <span class="cartcouponphrase">' + flavorText + '</span>\n') ########## Flavor text goes here ##########
        html.write('    </td>\n')
    for i in range(len(x)):
        html.write('    <td width="326" align="center" style="padding-right: 10px">\n')
        html.write('        <a href="javascript:eventApplyTime(' + str(i + 1) + ')"><img src="' + valueDict[x[i]] + '" width="100%" alt=""></a>\n')
        html.write('    </td>\n')
    html.write('</tr>\n')
    html.write('</table>\n')
    html.write('\n')
    html.write('<div align="center" style="padding: 15px 5px;">\n')
    html.write('<a href="' + tncEntryVar.get() + '" onClick="window.open("https://dp.image-gmkt.com/dp2016/SG/design/campaign/2020/03_Mar/0301/coupons/0301_2Coupon_TnCs.jpg","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=no, width=600,height=610,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">Terms and Conditions &#9656;</button></a>\n')
    html.write('<a href="https://www.qoo10.sg/gmkt.inc/Event/qchance.aspx" target="_blank"><button class="button buttonCpn">Get MameQ and Rewards &#9656;</button></a>\n')
    html.write('<a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/07/sendcoupon_WEB.html?2" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/07/sendcoupon_WEB.html?2","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=1, width=950,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">Send Coupon &#9656;</button></a>\n')
    html.write('<a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/11/howtousecoupon_WEB.html?2" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/11/howtousecoupon_WEB.html?2","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=1, width=950,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">How to use Coupon? &#9656;</button></a>\n')
    html.write('</div>\n')
    html.write('\n')
    html.write('\n')
    html.write('\n')
    ##### Mobile #####
    html.write('<style>\n')
    html.write('@import url("https://fonts.googleapis.com/css?family=Montserrat:400,500,700&display=swap");\n')
    html.write('.cartcoupontable {\n')
    html.write('    border: 1px solid #dbe1e2;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.cartcoupontable td{\n')
    html.write('    padding: 0px 5px;\n')
    html.write('}\n')
    html.write('\n')	
    html.write('.cartcoupondate {\n')
    html.write('    font-family: Montserrat, Helvetica, Arial, sans-serif;\n')
    html.write('    background: #ec2e3c;\n')
    html.write('    border-radius: 18px;\n')
    html.write('    border: 0px;\n')
    html.write('    color: #FFF;\n')
    html.write('    font-size: 25px;\n')
    html.write('    letter-spacing: 1px;\n')
    html.write('    padding: 7px 25px;\n')
    html.write('    text-align: center\n')
    html.write('}\n')
    html.write('\n')
    html.write('.cartcouponphrase {\n')
    html.write('    font-family: Montserrat, Helvetica, Arial, sans-serif;\n')
    html.write('    font-size: 35px;\n')
    html.write('    font-weight: 700;\n')
    html.write('    color: #333333;\n')
    html.write('    line-height: 27px;\n')
    html.write('}\n')
    html.write('\n')
    html.write('</style>\n')
    html.write('\n')
    html.write('<script>\n')
    html.write('function eventApplyTime(x) {\n')
    html.write('    var currentHour = (new Date()).getHours();\n')
    html.write('    switch(x) {\n')
    if len(x) >= 1:
        # Coupon 1 conditionals
        html.write('        case 1:\n')
        eidList = []
        timeList = []
        for eid, time in coupon1Dict.items():
            eidList.append(eid)
            timeList.append(time)
        if coupon1EID == 3:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '","' + str(eidList[2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '","' + str(timeList[2]) + '");\n')
        elif coupon1EID == 2:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '")\n')
        html.write('            break;\n')
        eidList.clear()
        timeList.clear()
    if len(x) >= 2:
        # Coupon 2 conditionals
        html.write('        case 2:\n')
        eid2List = []
        time2List = []
        for eid, time in coupon2Dict.items():
            eidList.append(eid)
            timeList.append(time)
        if coupon2EID == 3:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '","' + str(eidList[2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '","' + str(timeList[2]) + '");\n')
        elif coupon2EID == 2:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '")\n')
        html.write('            break;\n')
        eidList.clear()
        timeList.clear()
    if len(x) >= 3:
        # Coupon 3 conditionals
        html.write('        case 3:\n')
        eid3List = []
        time3List = []
        for eid, time in coupon3Dict.items():
            eidList.append(eid)
            timeList.append(time)
        print(eidList)
        print(timeList)
        if coupon3EID == 3:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '","' + str(eidList[2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '","' + str(timeList[2]) + '");\n')
        elif coupon3EID == 2:
            html.write('            var events = new Array("' + str(eidList[0]) + '","' + str(eidList[1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '","' + str(timeList[1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0]) + '")\n')
        html.write('            break;\n')
        eidList.clear()
        timeList.clear()
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
    html.write('<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#f8f8fa" class="cartcoupontable">\n')
    if len(flavorText) and len(date) != 0:
        html.write('    <tr>\n')
        html.write('        <td align="center" style="padding: 20px 0px;" width="50%" >\n')
        html.write('        <span class="cartcoupondate">' + date + '</span></td>\n')
        html.write('        <td width="50%" valign="middle" align="center">\n')
        html.write('            <span class="cartcouponphrase">' + flavorText + '</span></td>\n')
        html.write('    </tr>\n')
    html.write('    <tr>\n')
    html.write('        <td colspan="2" align="center" valign="top">\n')
    for i in range(len(x)):
        html.write('            <a href="javascript:eventApplyTime(' + str(i + 1) + ')"><img src="' + valueDict[x[i]] + '" width="42%" style="margin: 0px 10px;"></a>\n')
    html.write('        </td>\n')
    html.write('    </tr>\n')
    html.write('    <tr>\n')
    html.write('        <td colspan="3" align="center">&nbsp;</td>\n')
    html.write('    </tr>\n')
    html.write('</table>\n')
    html.write('<table width="100%" border="0">\n')
    html.write('<tr>\n')
    html.write('    <td><a href="' + tncEntryVar.get() + '" onClick="window.open("https://dp.image-gmkt.com/dp2016/SG/design/campaign/2020/03_Mar/0302/coupon/0302-0303_1Coupon_TnCs.jpg","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=no, width=600,height=700,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_01.png" width="100%"></a></td>\n')
    html.write('    <td><a href="https://www.qoo10.sg/gmkt.inc/Event/qchance.aspx" target="_blank"><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_02.png" width="100%"></a></td>\n')
    html.write('    <td><a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/sendcoupon_MB.html" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/sendcoupon_MB.html","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=no, width=600,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_03.png" width="100%"></a></td>\n')
    html.write('    <td><a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/howtousecoupon_MB.html" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/howtousecoupon_MB.html","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=no, width=600,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_04.png" width="100%"></a></td>\n')
    html.write('</tr>\n')
    html.write('</table>\n')

    html.close()
    messagebox.showinfo("Success!", 'Filename: "Coupon_' + date + '.html" Generated!')

######## GUI ########
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
frame1.grid(column=0, row=0, padx=5)

dateLabel = ttk.Label(frame1, text="Date")
dateLabel.grid(column=0, row=0, padx=10, sticky="nw")

dateEntryVar = tk.StringVar()
dateEntry = ttk.Entry(frame1, width=20, textvariable=dateEntryVar)
dateEntry.grid(column=1, row=0, pady=2, sticky="nw")

flavorLabel = ttk.Label(frame1, text="Flavor Text")
flavorLabel.grid(column=0, row=1, padx=10, sticky="nw")

flavorTextEntryVar = tk.StringVar()
flavorTextEntry = ttk.Entry(frame1, width=20, textvariable=flavorTextEntryVar)
flavorTextEntry.grid(column=1, row=1, pady=2, sticky="nw")

tncLabel = ttk.Label(frame1, text="T&Cs")
tncLabel.grid(column=0, row=2, padx=10, sticky="nw")

tncEntryVar = tk.StringVar()
tncEntry = ttk.Entry(frame1, width=20, textvariable=tncEntryVar)
tncEntry.grid(column=1, row=2, pady=2, sticky="nw")

## Frame 2 (Coupon 1 frame)
frame2 = ttk.Frame(win)
frame2.grid(column=1, row=0, padx=5)

coupon1Label = ttk.Label(frame2, text="Coupon 1")
coupon1Label.grid(column=0, row=0, columnspan=3)

coupon1ImgVar = tk.StringVar()
coupon1Img = ttk.Combobox(frame2, width=15, textvariable=coupon1ImgVar)
coupon1Img["values"] = valueDictKey
coupon1Img.grid(column=0, row=1, columnspan=3)

# Label for EID (Coupon 1)
c1EIDLabel = ttk.Label(frame2, text="EID")
c1EIDLabel.grid(column=0, row=2, columnspan=3)

# Entry for EID (Coupon 1)
c1e1Var = tk.StringVar()
c1e1Entry = ttk.Entry(frame2, width=5, textvariable=c1e1Var)
c1e1Entry.grid(column=0, row=3, padx=2, pady=2)

c1e2Var = tk.StringVar()
c1e2Entry = ttk.Entry(frame2, width=5, textvariable=c1e2Var)
c1e2Entry.grid(column=1, row=3, padx=2, pady=2)

c1e3Var = tk.StringVar()
c1e3Entry = ttk.Entry(frame2, width=5, textvariable=c1e3Var)
c1e3Entry.grid(column=2, row=3, padx=2, pady=2)

# Label for Timing (Coupon 1)
c1TimeLabel = ttk.Label(frame2, text="Hour (24hr Format)")
c1TimeLabel.grid(column=0, row=4, columnspan=3)

# Entry for timing (Coupon 1)
c1t1Var = tk.StringVar()
c1t1Entry = ttk.Entry(frame2, width=5, textvariable=c1t1Var)
c1t1Entry.grid(column=0, row=5, padx=2, pady=2)

c1t2Var = tk.StringVar()
c1t2Entry = ttk.Entry(frame2, width=5, textvariable=c1t2Var)
c1t2Entry.grid(column=1, row=5, padx=2, pady=2)

c1t3Var = tk.StringVar()
c1t3Entry = ttk.Entry(frame2, width=5, textvariable=c1t3Var)
c1t3Entry.grid(column=2, row=5, padx=2, pady=2)

## Frame 3 (Coupon 2 frame)
frame3 = ttk.Frame(win)
frame3.grid(column=2, row=0, padx=5)

coupon2Label = ttk.Label(frame3, text="Coupon 2")
coupon2Label.grid(column=0, row=0, columnspan=3)

coupon2ImgVar = tk.StringVar()
coupon2Img = ttk.Combobox(frame3, width=15, textvariable=coupon2ImgVar)
coupon2Img["values"] = valueDictKey
coupon2Img.grid(column=0, row=1, columnspan=3)

# Label for EID (Coupon 2)
c2EIDLabel = ttk.Label(frame3, text="EID")
c2EIDLabel.grid(column=0, row=2, columnspan=3)

# Entry for EID (Coupon 2)
c2e1Var = tk.StringVar()
c2e1Entry = ttk.Entry(frame3, width=5, textvariable=c2e1Var)
c2e1Entry.grid(column=0, row=3, padx=2, pady=2)

c2e2Var = tk.StringVar()
c2e2Entry = ttk.Entry(frame3, width=5, textvariable=c2e2Var)
c2e2Entry.grid(column=1, row=3, padx=2, pady=2)

c2e3Var = tk.StringVar()
c2e3Entry = ttk.Entry(frame3, width=5, textvariable=c2e3Var)
c2e3Entry.grid(column=2, row=3, padx=2, pady=2)

# Label for Timing (Coupon 2)
c2TimeLabel = ttk.Label(frame3, text="Hour (24hr Format)")
c2TimeLabel.grid(column=0, row=4, columnspan=3)

# Entry for timing (Coupon 2)
c2t1Var = tk.StringVar()
c2t1Entry = ttk.Entry(frame3, width=5, textvariable=c2t1Var)
c2t1Entry.grid(column=0, row=5, padx=2, pady=2)

c2t2Var = tk.StringVar()
c2t2Entry = ttk.Entry(frame3, width=5, textvariable=c2t2Var)
c2t2Entry.grid(column=1, row=5, padx=2, pady=2)

c2t3Var = tk.StringVar()
c2t3Entry = ttk.Entry(frame3, width=5, textvariable=c2t3Var)
c2t3Entry.grid(column=2, row=5, padx=2, pady=2)

## Frame 4 (Coupon 3 frame)
frame4 = ttk.Frame(win)
frame4.grid(column=3, row=0, padx=5)

coupon3Label = ttk.Label(frame4, text="Coupon 3")
coupon3Label.grid(column=0, row=0, columnspan=3)

coupon3ImgVar = tk.StringVar()
coupon3Img = ttk.Combobox(frame4, width=15, textvariable=coupon3ImgVar)
coupon3Img["values"] = valueDictKey
coupon3Img.grid(column=0, row=1, columnspan=3)

# Label for EID (Coupon 3)
c3EIDLabel = ttk.Label(frame4, text="EID")
c3EIDLabel.grid(column=0, row=2, columnspan=3)

# Entry for EID (Coupon 3)
c3e1Var = tk.StringVar()
c3e1Entry = ttk.Entry(frame4, width=5, textvariable=c3e1Var)
c3e1Entry.grid(column=0, row=3, padx=2, pady=2)

c3e2Var = tk.StringVar()
c3e2Entry = ttk.Entry(frame4, width=5, textvariable=c3e2Var)
c3e2Entry.grid(column=1, row=3, padx=2, pady=2)

c3e3Var = tk.StringVar()
c3e3Entry = ttk.Entry(frame4, width=5, textvariable=c3e3Var)
c3e3Entry.grid(column=2, row=3, padx=2, pady=2)

# Label for Timing (Coupon 3)
c3TimeLabel = ttk.Label(frame4, text="Hour (24hr Format)")
c3TimeLabel.grid(column=0, row=4, columnspan=3)

# Entry for timing (Coupon 3)
c3t1Var = tk.StringVar()
c3t1Entry = ttk.Entry(frame4, width=5, textvariable=c3t1Var)
c3t1Entry.grid(column=0, row=5)

c3t2Var = tk.StringVar()
c3t2Entry = ttk.Entry(frame4, width=5, textvariable=c3t2Var)
c3t2Entry.grid(column=1, row=5, padx=2, pady=2)

c3t3Var = tk.StringVar()
c3t3Entry = ttk.Entry(frame4, width=5, textvariable=c3t3Var)
c3t3Entry.grid(column=2, row=5, padx=2, pady=2)

## Frame 5 (Buttons frame)
frame5 = ttk.Frame(win)
frame5.grid(column=0, row=2, columnspan=4, sticky="nsew")

# Button
generateButton = ttk.Button(frame5, text="Generate", command=generate)
generateButton.grid(column=0, row=0)

## Labels
helpMsg = "WIP"
helpLabel = tk.Message(guide, text=helpMsg, width=500)
helpLabel.grid(row=0, column=0)

win.mainloop()