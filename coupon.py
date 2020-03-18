import re
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from collections import defaultdict
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import sys
import os

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

######## Google API client authorization ########
if __name__ == "__main__":
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(resource_path("coupon-generator.json"), scope)
    client = gspread.authorize(creds)
    sheet = client.open('Coupon Schedule')
    worksheet = sheet.worksheet("Coupon Summary")

######## Functions ########

## Populating design dropdown
def designs():
    global valueDict
    global valueDictKey
    valueDict = defaultdict(list) # Initialize collections dictionary of list
    valueDictKey = []
    designList = []

    for i in range(len(worksheet.col_values(1)) - 2):
        if len(worksheet.cell(i+3, 3).value) != 0: # Check if desktop design exists
            value = worksheet.cell(i+3,1).value
            valueDict[value].append(worksheet.cell(i+3,3).value) # Append desktop design URL to dictionary list
            valueDict[value].append(worksheet.cell(i+3,4).value) # Append mobile design URL to dictionary list

    for each in valueDict.keys():
        valueDictKey.append(each)

designs()

## Updating design entry box
def update(*args):
    desktopCoupon1Img.delete(first=0, last='end')
    desktopCoupon2Img.delete(first=0, last='end')
    desktopCoupon3Img.delete(first=0, last='end')
    mobileCoupon1Img.delete(first=0, last='end')
    mobileCoupon2Img.delete(first=0, last='end')
    mobileCoupon3Img.delete(first=0, last='end')
    try:
        if len(value1.get()) != 0:
            desktopCoupon1Img.insert(0, valueDict[value1Var.get()][0])
            mobileCoupon1Img.insert(0, valueDict[value1Var.get()][1])
        if len(value2.get()) != 0:
            desktopCoupon2Img.insert(0, valueDict[value2Var.get()][0])
            mobileCoupon2Img.insert(0, valueDict[value2Var.get()][1])
        if len(value3.get()) != 0:
            desktopCoupon3Img.insert(0, valueDict[value3Var.get()][0])
            mobileCoupon3Img.insert(0, valueDict[value3Var.get()][1])
    except:
        pass

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

    if len(desktopCoupon1ImgVar.get()) == 0:
        messagebox.showinfo("Oops!","At least 1 Coupon image is required!")
        return
    else:
        if len(desktopCoupon1ImgVar.get()) != 0: 
            noCoupon.append((desktopCoupon1ImgVar.get(), mobileCoupon1ImgVar.get()))
        if len(desktopCoupon2ImgVar.get()) != 0:
            noCoupon.append((desktopCoupon2ImgVar.get(), mobileCoupon2ImgVar.get()))   
        if len(desktopCoupon3ImgVar.get()) != 0:
            noCoupon.append((desktopCoupon3ImgVar.get(), mobileCoupon3ImgVar.get()))
    date = dateEntryVar.get()
    flavorText = flavorTextEntryVar.get()
    print(len(noCoupon))
    if len(noCoupon) >= 1:
        if value1Var.get() in valueDictKey:
            pass
        else:
            worksheet.append_row([value1Var.get(),'', desktopCoupon1ImgVar.get(), mobileCoupon1ImgVar.get()])
        # Coupon 1 variables
        if len(c1e1Var.get()) == 0:
            messagebox.showinfo("Error", "You require at least 1 EID for the first coupon!")
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
        if value2Var.get() in valueDictKey:
            pass
        else:
            if len(desktopCoupon2ImgVar.get()) == 0:
                messagebox.showinfo("Oops!","You require an image for the second coupon!")
                return
            else:
                worksheet.append_row([value2Var.get(),'', desktopCoupon2ImgVar.get(), mobileCoupon2ImgVar.get()])
        # Coupon 2 variables
        if len(c2e1Var.get()) == 0:
            messagebox.showinfo("Error", "You require at least 1 EID for the second coupon!")
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
        if value3Var.get() in valueDictKey:
            pass
        else:
            if len(desktopCoupon3ImgVar.get()) == 0:
                messagebox.showinfo("Oops!","You require an image for the third coupon!")
                return
            else:
                worksheet.append_row([value3Var.get(),'', desktopCoupon3ImgVar.get(), mobileCoupon3ImgVar.get()])
        # Coupon 3 variables
        if len(c3e1Var.get()) == 0:
            messagebox.showinfo("Error", "You require at least 1 EID for third coupon!")
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
    html.write('<!-- Desktop -->\n')
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
    html.write('    font-family: Montserrat, Helvetica, Arial, sans-serif;\n')
    html.write('    background: #ec2e3c;\n')
    html.write('    border-radius: 18px;\n')
    html.write('    border: 0px;\n')
    html.write('    color: #FFF;\n')
    html.write('    font-size: 16px;\n')
    html.write('    letter-spacing: 1px;\n')
    html.write('    padding: 7px 25px;\n')
    html.write('}\n')
    html.write('.cartcouponphrase {\n')
    html.write('    font-family: Montserrat, Helvetica, Arial, sans-serif;\n')
    html.write('    font-size: 38px;\n')
    html.write('    font-weight: 700;\n')
    html.write('    color: #333333;\n')
    html.write('    line-height: 40px;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.tncModal {\n')
    html.write('    display: none; /* Hidden by default */\n')
    html.write('    position: fixed; /* Stay in place */\n')
    html.write('    z-index: 1; /* Sit on top */\n')
    html.write('    padding-top: 100px; /* Location of the box */\n')
    html.write('    left: 0;\n')
    html.write('    top: 0;\n')
    html.write('    width: 100%; /* Full width */\n')
    html.write('    height: 100%; /* Full height */\n')
    html.write('    overflow: auto; /* Enable scroll if needed */\n')
    html.write('    background-color: rgb(0,0,0); /* Fallback color */\n')
    html.write('    background-color: rgba(0,0,0,0.4); /* Black w/ opacity */\n')
    html.write('}\n')
    html.write('\n')
    html.write('.tncModalContent {\n')
    html.write('    background-color: #fefefe;\n')
    html.write('    padding: 20px;\n')
    html.write('    border: 1px solid #888;\n')
    html.write('    width: 80%;\n')
    html.write('    width: 750px;\n')
    html.write('    text-align: left;\n')
    html.write('    font-size: 14px;\n')
    html.write('    font-family: arial, sans-serif;\n')
    html.write('    border-radius: 25px;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.tncHeader {\n')
    html.write('    font-size: 25;\n')
    html.write('    font-weight: bold;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.buttonDisplayTopright {\n')
    html.write('    float: right;\n')
    html.write('    color: #aaa;\n')
    html.write('    font-size: 21px;\n')
    html.write('    font-weight: bold;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.buttonDisplayTopright:hover,\n')
    html.write('.buttonDisplayTopright:focus {\n')
    html.write('    color: #000;\n')
    html.write('    text-decoration: none;\n')
    html.write('    cursor: pointer;\n')
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
        eidList = [[],[],[]] # Nested list to store all time variables
        timeList = [[],[],[]] # Nested list to store all time variables
        for eid, time in coupon1Dict.items():
            eidList[0].append(eid)
            timeList[0].append(time)
        if coupon1EID == 3:
            html.write('            var events = new Array("' + str(eidList[0][0]) + '","' + str(eidList[0][1]) + '","' + str(eidList[0][2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0][0]) + '","' + str(timeList[0][1]) + '","' + str(timeList[0][2]) + '");\n')
        elif coupon1EID == 2:
            html.write('            var events = new Array("' + str(eidList[0][0]) + '","' + str(eidList[0][1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0][0]) + '","' + str(timeList[0][1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0][0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0][0]) + '")\n')
        html.write('            break;\n')
    if len(x) >= 2:
        # Coupon 2 conditionals
        html.write('        case 2:\n')
        for eid, time in coupon2Dict.items():
            eidList[1].append(eid)
            timeList[1].append(time)
        if coupon2EID == 3:
            html.write('            var events = new Array("' + str(eidList[1][0]) + '","' + str(eidList[1][1]) + '","' + str(eidList[1][2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[1][0]) + '","' + str(timeList[1][1]) + '","' + str(timeList[1][2]) + '");\n')
        elif coupon2EID == 2:
            html.write('            var events = new Array("' + str(eidList[1][0]) + '","' + str(eidList[1][1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[1][0]) + '","' + str(timeList[1][1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[1][0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[1][0]) + '")\n')
        html.write('            break;\n')
    if len(x) >= 3:
        # Coupon 3 conditionals
        html.write('        case 3:\n')
        for eid, time in coupon3Dict.items():
            eidList[2].append(eid)
            timeList[2].append(time)
        if coupon3EID == 3:
            html.write('            var events = new Array("' + str(eidList[2][0]) + '","' + str(eidList[2][1]) + '","' + str(eidList[2][2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[2][0]) + '","' + str(timeList[2][1]) + '","' + str(timeList[2][2]) + '");\n')
        elif coupon3EID == 2:
            html.write('            var events = new Array("' + str(eidList[2][0]) + '","' + str(eidList[2][1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[2][0]) + '","' + str(timeList[2][1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[2][0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[2][0]) + '")\n')
        html.write('            break;\n')
    html.write('    }\n')
    html.write('    if (hours.length == 1) {Util.EventApply(events[0]);}\n')
    html.write('        else if (hours.length == 2) {\n')
    html.write('            else if (currentHour >= hours[1]) {Util.EventApply(events[1]);}\n')
    html.write('        }\n')
    html.write('        else if (hours.length == 3) {\n')
    html.write('            if (currentHour >= hours[0] && currentHour < hours[1]) {Util.EventApply(events[0]);}\n')
    html.write('                else if (currentHour >= hours[1] && currentHour < hours[2]) {Util.EventApply(events[1]);}\n')
    html.write('                else if (currentHour >= hours[2]) {Util.EventApply(events[2]);}\n')
    html.write('    }\n')
    html.write('};\n')
    html.write('</script>\n')
    html.write('\n')
    if len(x) == 3:
        html.write('    <table width="980" border="0" align="center" cellpadding="0" cellspacing="0">\n')
        html.write('        <tr>\n')
        html.write('            <td><a href="javascript:eventApplyTime(1)"><img src="' + x[0][0] + '" width="327"></a></td>\n')
        html.write('            <td><a href="javascript:eventApplyTime(2)"><img src="' + x[1][0] + '" width="326"></a></td>\n')
        html.write('            <td><a href="javascript:eventApplyTime(3)"><img src="' + x[2][0] + '" width="327"></a></td>\n')
        html.write('        </tr>\n')
        html.write('    </table>\n')
        html.write('\n')
    else:
        html.write('<table width="980" height="200" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#f8f8fa" class="cartcoupontable">\n')
        html.write('<tr>\n')
        if len(flavorText) and len(date) != 0:
            if len(x) == 1:
                html.write('    <td width="654" style="padding-left: 30px;">\n')
            else:
                html.write('    <td width="328" style="padding-left: 30px;">\n')
            html.write('        <span class="cartcoupondate">' + date + '</span><br><br>\n') ########## Flavor date goes here ##########
            html.write('        <span class="cartcouponphrase">' + flavorText + '</span>\n') ########## Flavor text goes here ##########
            html.write('    </td>\n')
        for i in range(len(x)):
            html.write('    <td width="326" align="center" style="padding-right: 15px; padding-top:10px; padding-bottom: 10px;">\n')
            html.write('        <a href="javascript:eventApplyTime(' + str(i + 1) + ')"><img src="' + x[i][0] + '" width="100%" alt=""></a>\n')
            html.write('    </td>\n')
        html.write('</tr>\n')
        html.write('</table>\n')
        html.write('\n')
    html.write('<div align="center" style="padding: 15px 5px;">\n')
    html.write('    <div>\n')
    html.write('        <button class="button buttonCpn" onclick="document.getElementById(\'tnc\').style.display=\'block\'">Terms and Conditions &#9656;</button></a>\n')
    html.write('        <a href="https://www.qoo10.sg/gmkt.inc/Event/qchance.aspx" target="_blank"><button class="button buttonCpn">Get MameQ and Rewards &#9656;</button></a>\n')
    html.write('        <a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/07/sendcoupon_WEB.html?2" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/07/sendcoupon_WEB.html?2","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=1, width=950,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">Send Coupon &#9656;</button></a>\n')
    html.write('        <a href="https://dp.image-gmkt.com/dp2016/SG/design/How_to_use_coupons.jpg" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/11/howtousecoupon_WEB.html?2","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=1, width=950,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><button class="button buttonCpn">How to use Coupon? &#9656;</button></a>\n')
    html.write('    </div>\n')
    html.write('\n')
    html.write('    <div id="tnc" class="tncModal">\n')
    html.write('        <div class="tncModalContent">\n')
    html.write('            <span onclick="document.getElementById(\'tnc\').style.display=\'none\'" class="buttonDisplayTopright">&times;</span>\n')
    html.write('            <p class="tncHeader">Terms and Conditions</p>\n')
    html.write('            <ul>\n')
    html.write('                <li>- Event runs from <b>' + date + ' (Event Period)</b>.</li>\n')
    html.write('                <li>- Coupons may only be received during <u>Event Period</u>.</li>\n')
    html.write('                <li>- Coupons are valid only for the duration of the <u>Event Period</u>.</li>\n')
    value = []
    altValue = []
    if re.search("[0-9]/[0-9]", value1Var.get()):
        value.append(value1Var.get().split('/'))
        altValue.append("")
    else:
        altValue.append(value1Var.get())
    if re.search("[0-9]/[0-9]", value2Var.get()):
        value.append(value2Var.get().split('/'))
        altValue.append("")
    else:
        altValue.append(value2Var.get())
    if re.search("[0-9]/[0-9]", value3Var.get()):
        value.append(value3Var.get().split('/'))
        altValue.append("")
    else:
        altValue.append(value3Var.get())
    mmq = [mmq1EntryVar.get(), mmq2EntryVar.get(), mmq3EntryVar.get()]
    qty = [qty1EntryVar.get(), qty2EntryVar.get(), qty3EntryVar.get()]
    for i in range(len(x)):
        if len(altValue[i]) == 0:
            html.write('               <li>- <b>' + value[i][0] + ' cart coupon</b> is applicable towards valid purchases with a <b>minimum order value of $' + value[i][1] + '</b>.</li>\n')
        else:
            html.write('               <li>- <b>' + altValue[i] + '</b>.</li>\n')
        html.write('            <ul>\n')
        html.write('                <li>&nbsp&nbsp&nbsp&nbsp- Coupon redemption available at daily at:.</li>\n')
        html.write('                    <ul>\n')
        for eachTime in range(len(timeList[i])):
            html.write('                        <li>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp- ' + str(timeList[i][eachTime]) + ':00</li>\n')
        html.write('                    </ul>\n')
        html.write('                <li>&nbsp&nbsp&nbsp&nbsp- <b>' + str(mmq[i]) + ' MameQ</b> is required for this cart coupon.</li>\n')
        html.write('                <li>&nbsp&nbsp&nbsp&nbsp-  Coupon is limited to a total of <b>' + str(qty[i]) + '</b> applicants daily.</li>\n')
        html.write('            </ul>')
        html.write('            <br>\n')
    html.write('                <li>- Applicants may only once per coupon during the <u>Event Period</u>.</li>\n')
    html.write('                <li>- Event application/purchases are only available within Qoo10 Singapore (www.qoo10.sg).</li>\n')
    html.write('                <li>- Entries received after the <u>Event Period</u> will be invalid and no refunds will be issued.</li>\n')
    html.write('                <li>- Coupons cannot be sold or exchanged for cash.</li>\n')
    html.write('                <li>- Qoo10 Singapore reserves the rights to make any amendments to the Terms and Conditions herein at any point in time, without prior notice.</li>\n')
    html.write('                <li>- By participating in this event, you agree to be bound by the Terms and Conditions, the User Agreement, and the decisions of Qoo10.</li>\n')
    html.write('            </ul>\n')
    html.write('        </div>\n')
    html.write('    </div>\n')
    html.write('</div>\n')
    html.write('\n')
    html.write('\n')
    html.write('\n')
    html.write('\n')
    html.write('\n')
    html.write('\n')
    html.write('\n')
    html.write('\n')
    ##### Mobile #####
    html.write('<!-- Mobile -->\n')
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
    html.write('    font-size: 15px;\n')
    html.write('    letter-spacing: 1px;\n')
    html.write('    padding: 7px 25px;\n')
    html.write('    text-align: center\n')
    html.write('}\n')
    html.write('\n')
    html.write('.cartcouponphrase {\n')
    html.write('    font-family: Montserrat, Helvetica, Arial, sans-serif;\n')
    html.write('    font-size: 25px;\n')
    html.write('    font-weight: 700;\n')
    html.write('    color: #333333;\n')
    html.write('    line-height: 27px;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.tncModal {\n')
    html.write('    display: none; /* Hidden by default */\n')
    html.write('    position: fixed; /* Stay in place */\n')
    html.write('    z-index: 1; /* Sit on top */\n')
    html.write('    padding-top: 100px; /* Location of the box */\n')
    html.write('    left: 0;\n')
    html.write('    top: 0;\n')
    html.write('    width: 100%; /* Full width */\n')
    html.write('    height: 100%; /* Full height */\n')
    html.write('    overflow: auto; /* Enable scroll if needed */\n')
    html.write('    background-color: rgb(0,0,0); /* Fallback color */\n')
    html.write('    background-color: rgba(0,0,0,0.4); /* Black w/ opacity */\n')
    html.write('}\n')
    html.write('\n')
    html.write('.tncModalContent {\n')
    html.write('    background-color: #fefefe;\n')
    html.write('    padding: 20px;\n')
    html.write('    border: 1px solid #888;\n')
    html.write('    width: 80%;\n')
    html.write('    width: 750px;\n')
    html.write('    text-align: left;\n')
    html.write('    font-size: 14px;\n')
    html.write('    font-family: arial, sans-serif;\n')
    html.write('    border-radius: 25px;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.tncHeader {\n')
    html.write('    font-size: 25;\n')
    html.write('    font-weight: bold;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.buttonDisplayTopright {\n')
    html.write('    float: right;\n')
    html.write('    color: #aaa;\n')
    html.write('    font-size: 21px;\n')
    html.write('    font-weight: bold;\n')
    html.write('}\n')
    html.write('\n')
    html.write('.buttonDisplayTopright:hover,\n')
    html.write('.buttonDisplayTopright:focus {\n')
    html.write('    color: #000;\n')
    html.write('    text-decoration: none;\n')
    html.write('    cursor: pointer;\n')
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
        if coupon1EID == 3:
            html.write('            var events = new Array("' + str(eidList[0][0]) + '","' + str(eidList[0][1]) + '","' + str(eidList[0][2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0][0]) + '","' + str(timeList[0][1]) + '","' + str(timeList[0][2]) + '");\n')
        elif coupon1EID == 2:
            html.write('            var events = new Array("' + str(eidList[0][0]) + '","' + str(eidList[0][1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0][0]) + '","' + str(timeList[0][1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[0][0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[0][0]) + '")\n')
        html.write('            break;\n')
    if len(x) >= 2:
        # Coupon 2 conditionals
        html.write('        case 2:\n')
        if coupon2EID == 3:
            html.write('            var events = new Array("' + str(eidList[1][0]) + '","' + str(eidList[1][1]) + '","' + str(eidList[1][2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[1][0]) + '","' + str(timeList[1][1]) + '","' + str(timeList[1][2]) + '");\n')
        elif coupon2EID == 2:
            html.write('            var events = new Array("' + str(eidList[1][0]) + '","' + str(eidList[1][1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[1][0]) + '","' + str(timeList[1][1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[1][0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[1][0]) + '")\n')
        html.write('            break;\n')
    if len(x) >= 3:
        # Coupon 3 conditionals
        html.write('        case 3:\n')
        if coupon3EID == 3:
            html.write('            var events = new Array("' + str(eidList[2][0]) + '","' + str(eidList[2][1]) + '","' + str(eidList[2][2]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[2][0]) + '","' + str(timeList[2][1]) + '","' + str(timeList[2][2]) + '");\n')
        elif coupon3EID == 2:
            html.write('            var events = new Array("' + str(eidList[2][0]) + '","' + str(eidList[2][1]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[2][0]) + '","' + str(timeList[2][1]) + '");\n')
        else:
            html.write('            var events = new Array("' + str(eidList[2][0]) + '")\n')
            html.write('            var hours = new Array("' + str(timeList[2][0]) + '")\n')
        html.write('            break;\n')
    html.write('    }\n')
    html.write('    if (hours.length == 1) {Util.EventApply(events[0]);}\n')
    html.write('        else if (hours.length == 2) {\n')
    html.write('            else if (currentHour >= hours[1]) {Util.EventApply(events[1]);}\n')
    html.write('        }\n')
    html.write('        else if (hours.length == 3) {\n')
    html.write('            if (currentHour >= hours[0] && currentHour < hours[1]) {Util.EventApply(events[0]);}\n')
    html.write('                else if (currentHour >= hours[1] && currentHour < hours[2]) {Util.EventApply(events[1]);}\n')
    html.write('                else if (currentHour >= hours[2]) {Util.EventApply(events[2]);}\n')
    html.write('    }\n')
    html.write('};\n')
    html.write('</script>\n')
    html.write('\n')
    if len(x) == 3:
        html.write('<table width="100%" border="0" cellpadding="0" cellspacing="0">\n')
        html.write('    <tr>')
        html.write('        <td width="50%"><img src="' + flavorText + '" width="100%"></td>\n')
        html.write('        <td width="50%" bgcolor="#ffe7e3"><a href="javascript:eventApplyTime(1)"><img src="' + x[0][1] + '" width="100%" alt=""></a></td>\n')
        html.write('    </tr>\n')
        html.write('    <tr>\n')
        html.write('        <td width="50%" bgcolor="#ffe7e3"><a href="javascript:eventApplyTime(2)"><img src="' + x[1][1] + '" width="100%" alt=""></a></td>\n')
        html.write('        <td width="50%" bgcolor="#ffe7e3"><a href="javascript:eventApplyTime(3)"><img src="' + x[2][1] + '" width="100%" alt=""></a></td>\n')
        html.write('    </tr>\n')
        html.write('</table>\n')
    else:
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
            html.write('            <a href="javascript:eventApplyTime(' + str(i + 1) + ')"><img src="' + x[i][1] + '" width="42%" style="margin: 0px 10px;"></a>\n')
        html.write('        </td>\n')
        html.write('    </tr>\n')
        html.write('    <tr>\n')
        html.write('        <td colspan="3" align="center">&nbsp;</td>\n')
        html.write('    </tr>\n')
        html.write('</table>\n')
    html.write('<table width="100%" border="0">\n')
    html.write('    <tr>\n')
    html.write('        <td><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_01.png" width="100%" onclick="document.getElementById(\'tnc\').style.display=\'block\'"></td>\n')
    html.write('        <td><a href="https://www.qoo10.sg/gmkt.inc/Event/qchance.aspx" target="_blank"><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_02.png" width="100%"></a></td>\n')
    html.write('        <td><a href="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/sendcoupon_MB.html" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/sendcoupon_MB.html","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=no, width=600,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_03.png" width="100%"></a></td>\n')
    html.write('        <td><a href="https://dp.image-gmkt.com/dp2016/SG/design/How_to_use_coupons.jpg" onClick="window.open("http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/howtousecoupon_MB.html","window","location=no, directories=no,resizable=no,status=no,toolbar=no,menubar=no, width=600,height=800,left=300, top=20, scrollbars=no");return false" onFocus="this.blur()"/><img src="http://dp.image-gmkt.com/dp2016/SG/design/PM1/2019/12/buttons_04.png" width="100%"></a></td>\n')
    html.write('    </tr>\n')
    html.write('</table>\n')
    html.write('<div id="tnc" class="tncModal">\n')
    html.write('    <div class="tncModalContent">\n')
    html.write('        <span onclick="document.getElementById(\'tnc\').style.display=\'none\'" class="buttonDisplayTopright">&times;</span>\n')
    html.write('        <p class="tncHeader">Terms and Conditions</p>\n')
    html.write('        <ul>\n')
    html.write('            <li>- Event runs from <b>' + date + ' (Event Period)</b>.</li>\n')
    html.write('            <li>- Coupons may only be received during <u>Event Period</u>.</li>\n')
    html.write('            <li>- Coupons are valid only for the duration of the <u>Event Period</u>.</li>\n')
    for i in range(len(x)):
        if len(altValue[i]) == 0:
            html.write('            <li>- <b>' + value[i][0] + ' cart coupon</b> is applicable towards valid purchases with a <b>minimum order value of $' + value[i][1] + '</b>.</li>\n')
        else:
            html.write('            <li>- <b>' + altValue[i] + '</b>.</li>\n')
        html.write('            <ul>\n')
        html.write('                <li>&nbsp&nbsp&nbsp&nbsp- Coupon redemption available at daily at:.</li>\n')
        html.write('                    <ul>\n')
        for eachTime in range(len(timeList[i])):
            html.write('                        <li>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp- ' + str(timeList[i][eachTime]) + ':00</li>\n')
        html.write('                    </ul>\n')
        html.write('                <li>&nbsp&nbsp&nbsp&nbsp- <b>' + str(mmq[i]) + ' MameQ</b> is required for this cart coupon.</li>\n')
        html.write('                <li>&nbsp&nbsp&nbsp&nbsp- Coupon is limited to a total of <b>' + str(qty[i]) + '</b> applicants daily.</li>\n')
        html.write('            </ul>')
        html.write('            <br>\n')
    html.write('            <li>- Applicants may only once per coupon during the <u>Event Period</u>.</li>\n')
    html.write('            <li>- Event application/purchases are only available within Qoo10 Singapore (www.qoo10.sg).</li>\n')
    html.write('            <li>- Entries received after the <u>Event Period</u> will be invalid and no refunds will be issued.</li>\n')
    html.write('            <li>- Coupons cannot be sold or exchanged for cash.</li>\n')
    html.write('            <li>- Qoo10 Singapore reserves the rights to make any amendments to the Terms and Conditions herein at any point in time, without prior notice.</li>\n')
    html.write('            <li>- By participating in this event, you agree to be bound by the Terms and Conditions, the User Agreement, and the decisions of Qoo10.</li>\n')
    html.write('        </ul>\n')
    html.write('    </div>\n')
    html.write('</div>\n')
    html.close()
    messagebox.showinfo("Success!", 'Filename: "Coupon_' + date + "_" + flavorText + '.html" Generated!')

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
frame1 = tk.Frame(win)
frame1.grid(column=0, row=0, padx=5)

dateLabel = ttk.Label(frame1, text="Date")
dateLabel.grid(column=0, row=1, padx=10, sticky="nw")

dateEntryVar = tk.StringVar()
dateEntry = ttk.Entry(frame1, width=20, textvariable=dateEntryVar)
dateEntry.grid(column=1, row=1, pady=2, sticky="nw")

flavorLabel = ttk.Label(frame1, text="Flavor Text / Design")
flavorLabel.grid(column=0, row=2, padx=10, sticky="nw")

flavorTextEntryVar = tk.StringVar()
flavorTextEntry = ttk.Entry(frame1, width=20, textvariable=flavorTextEntryVar)
flavorTextEntry.grid(column=1, row=2, pady=2, sticky="nw")

## Frame 2 (Coupon 1 frame)
frame2 = tk.Frame(win, highlightbackground="black", highlightthickness=1, padx=2, pady=2)
frame2.grid(column=1, row=0, padx=5, pady=5)

coupon1Label = ttk.Label(frame2, text="Coupon 1")
coupon1Label.grid(column=0, row=0, columnspan=3)

# Value
value1Label = ttk.Label(frame2, text="Value")
value1Label.grid(column=0, row=1)

value1Var = tk.StringVar()
value1 = ttk.Combobox(frame2, width=8, textvariable=value1Var)
value1["values"] = valueDictKey
value1.grid(column=1, row=1, pady=2, columnspan=2)
value1Var.trace("w", update)

# Design
design1Label = ttk.Label(frame2, text="Design")
design1Label.grid(column=0, row=2, columnspan=3)

desktopDesign1Label = ttk.Label(frame2, text="Desktop")
desktopDesign1Label.grid(column=0, row=3)

desktopCoupon1ImgVar = tk.StringVar()
desktopCoupon1Img = ttk.Entry(frame2, width=11, textvariable=desktopCoupon1ImgVar)
desktopCoupon1Img.grid(column=1, row=3, pady=2, columnspan=2)

mobileDesign1Label = ttk.Label(frame2, text="Mobile")
mobileDesign1Label.grid(column=0, row=4)

mobileCoupon1ImgVar = tk.StringVar()
mobileCoupon1Img = ttk.Entry(frame2, width=11, textvariable=mobileCoupon1ImgVar)
mobileCoupon1Img.grid(column=1, row=4, pady=2, columnspan=2)

# Qty
qty1Label = ttk.Label(frame2, text="Qty")
qty1Label.grid(column=0, row=5)

# Entry for Qty
qty1EntryVar = tk.StringVar()
qty1 = ttk.Entry(frame2, width=11, textvariable=qty1EntryVar)
qty1.insert(0, "0")
qty1.grid(column=1, row=5, padx=2, pady=2, columnspan=2)

# MameQ
mmq1Label = ttk.Label(frame2, text="No. MameQ")
mmq1Label.grid(column=0, row=6, columnspan=2)

# Entry for MMQ
mmq1EntryVar = tk.StringVar()
mmq1 = ttk.Entry(frame2, width=4, textvariable=mmq1EntryVar)
mmq1.insert(0, "0")
mmq1.grid(column=2, row=6, padx=2, pady=2)

# Label for EID (Coupon 1)
c1EIDLabel = ttk.Label(frame2, text="Encrypted EID")
c1EIDLabel.grid(column=0, row=7, columnspan=3)

# Entry for EID (Coupon 1)
c1e1Var = tk.StringVar()
c1e1Entry = ttk.Entry(frame2, width=5, textvariable=c1e1Var)
c1e1Entry.grid(column=0, row=8, padx=2, pady=2)

c1e2Var = tk.StringVar()
c1e2Entry = ttk.Entry(frame2, width=5, textvariable=c1e2Var)
c1e2Entry.grid(column=1, row=8, padx=2, pady=2)

c1e3Var = tk.StringVar()
c1e3Entry = ttk.Entry(frame2, width=5, textvariable=c1e3Var)
c1e3Entry.grid(column=2, row=8, padx=2, pady=2)

# Label for Timing (Coupon 1)
c1TimeLabel = ttk.Label(frame2, text="Hour (24hr Format)")
c1TimeLabel.grid(column=0, row=9, columnspan=3)

# Entry for timing (Coupon 1)
c1t1Var = tk.StringVar()
c1t1Entry = ttk.Entry(frame2, width=5, textvariable=c1t1Var)
c1t1Entry.grid(column=0, row=10, padx=2, pady=2)

c1t2Var = tk.StringVar()
c1t2Entry = ttk.Entry(frame2, width=5, textvariable=c1t2Var)
c1t2Entry.grid(column=1, row=10, padx=2, pady=2)

c1t3Var = tk.StringVar()
c1t3Entry = ttk.Entry(frame2, width=5, textvariable=c1t3Var)
c1t3Entry.grid(column=2, row=10, padx=2, pady=2)

## Frame 3 (Coupon 2 frame)
frame3 = tk.Frame(win, highlightbackground="black", highlightthickness=1, padx=2, pady=2)
frame3.grid(column=2, row=0, padx=5, pady=5)

# Coupon 2 Label
coupon2Label = ttk.Label(frame3, text="Coupon 2")
coupon2Label.grid(column=0, row=0, columnspan=3)

# Value
value2Label = ttk.Label(frame3, text="Value")
value2Label.grid(column=0, row=1)

value2Var = tk.StringVar()
value2 = ttk.Combobox(frame3, width=8, textvariable=value2Var)
value2["values"] = valueDictKey
value2.grid(column=1, row=1, pady=2, columnspan=2)
value2Var.trace("w", update)

# Design
design2Label = ttk.Label(frame3, text="Design")
design2Label.grid(column=0, row=2, columnspan=3)

desktopDesign2Label = ttk.Label(frame3, text="Desktop")
desktopDesign2Label.grid(column=0, row=3)

desktopCoupon2ImgVar = tk.StringVar()
desktopCoupon2Img = ttk.Entry(frame3, width=11, textvariable=desktopCoupon2ImgVar)
desktopCoupon2Img.grid(column=1, row=3, pady=2, columnspan=2)

mobileDesign2Label = ttk.Label(frame3, text="Mobile")
mobileDesign2Label.grid(column=0, row=4)

mobileCoupon2ImgVar = tk.StringVar()
mobileCoupon2Img = ttk.Entry(frame3, width=11, textvariable=mobileCoupon2ImgVar)
mobileCoupon2Img.grid(column=1, row=4, pady=2, columnspan=2)

# Qty
qty2Label = ttk.Label(frame3, text="Qty")
qty2Label.grid(column=0, row=5)

# Entry for Qty
qty2EntryVar = tk.StringVar()
qty2 = ttk.Entry(frame3, width=11, textvariable=qty2EntryVar)
qty2.insert(0, "0")
qty2.grid(column=1, row=5, padx=2, pady=2, columnspan=2)

# MameQ
mmq2Label = ttk.Label(frame3, text="No. MameQ")
mmq2Label.grid(column=0, row=6, columnspan=2)

# Entry for MMQ
mmq2EntryVar = tk.StringVar()
mmq2 = ttk.Entry(frame3, width=4, textvariable=mmq2EntryVar)
mmq2.insert(0, "0")
mmq2.grid(column=2, row=6, padx=2, pady=2)

# Label for EID (Coupon 2)
c2EIDLabel = ttk.Label(frame3, text="Encrypted EID")
c2EIDLabel.grid(column=0, row=7, columnspan=3)

# Entry for EID (Coupon 2)
c2e1Var = tk.StringVar()
c2e1Entry = ttk.Entry(frame3, width=5, textvariable=c2e1Var)
c2e1Entry.grid(column=0, row=8, padx=2, pady=2)

c2e2Var = tk.StringVar()
c2e2Entry = ttk.Entry(frame3, width=5, textvariable=c2e2Var)
c2e2Entry.grid(column=1, row=8, padx=2, pady=2)

c2e3Var = tk.StringVar()
c2e3Entry = ttk.Entry(frame3, width=5, textvariable=c2e3Var)
c2e3Entry.grid(column=2, row=8, padx=2, pady=2)

# Label for Timing (Coupon 2)
c2TimeLabel = ttk.Label(frame3, text="Hour (24hr Format)")
c2TimeLabel.grid(column=0, row=9, columnspan=3)

# Entry for timing (Coupon 2)
c2t1Var = tk.StringVar()
c2t1Entry = ttk.Entry(frame3, width=5, textvariable=c2t1Var)
c2t1Entry.grid(column=0, row=10, padx=2, pady=2)

c2t2Var = tk.StringVar()
c2t2Entry = ttk.Entry(frame3, width=5, textvariable=c2t2Var)
c2t2Entry.grid(column=1, row=10, padx=2, pady=2)

c2t3Var = tk.StringVar()
c2t3Entry = ttk.Entry(frame3, width=5, textvariable=c2t3Var)
c2t3Entry.grid(column=2, row=10, padx=2, pady=2)

## Frame 4 (Coupon 3 frame)
frame4 = tk.Frame(win, highlightbackground="black", highlightthickness=1, padx=2, pady=2)
frame4.grid(column=3, row=0, padx=5, pady=5)

coupon3Label = ttk.Label(frame4, text="Coupon 3")
coupon3Label.grid(column=0, row=0, columnspan=3)

# Value
value3Label = ttk.Label(frame4, text="Value")
value3Label.grid(column=0, row=1)

value3Var = tk.StringVar()
value3 = ttk.Combobox(frame4, width=8, textvariable=value3Var)
value3["values"] = valueDictKey
value3.grid(column=1, row=1, pady=2, columnspan=2)
value3Var.trace("w", update)

# Design
design3Label = ttk.Label(frame4, text="Design")
design3Label.grid(column=0, row=2, columnspan=3)

desktopDesign3Label = ttk.Label(frame4, text="Desktop")
desktopDesign3Label.grid(column=0, row=3)

desktopCoupon3ImgVar = tk.StringVar()
desktopCoupon3Img = ttk.Entry(frame4, width=11, textvariable=desktopCoupon3ImgVar)
desktopCoupon3Img.grid(column=1, row=3, pady=2, columnspan=2)

mobileDesign3Label = ttk.Label(frame4, text="Mobile")
mobileDesign3Label.grid(column=0, row=4)

mobileCoupon3ImgVar = tk.StringVar()
mobileCoupon3Img = ttk.Entry(frame4, width=11, textvariable=mobileCoupon3ImgVar)
mobileCoupon3Img.grid(column=1, row=4, pady=2, columnspan=2)

# Qty
qty3Label = ttk.Label(frame4, text="Qty")
qty3Label.grid(column=0, row=5)

# Entry for Qty
qty3EntryVar = tk.StringVar()
qty3 = ttk.Entry(frame4, width=11, textvariable=qty3EntryVar)
qty3.insert(0, "0")
qty3.grid(column=1, row=5, padx=2, pady=2, columnspan=2)

# MameQ
mmq3Label = ttk.Label(frame4, text="No. MameQ")
mmq3Label.grid(column=0, row=6, columnspan=2)

# Entry for MMQ
mmq3EntryVar = tk.StringVar()
mmq3 = ttk.Entry(frame4, width=4, textvariable=mmq3EntryVar)
mmq3.insert(0, "0")
mmq3.grid(column=2, row=6, padx=2, pady=2)

# Label for EID (Coupon 3)
c3EIDLabel = ttk.Label(frame4, text="Encrypted EID")
c3EIDLabel.grid(column=0, row=7, columnspan=3)

# Entry for EID (Coupon 3)
c3e1Var = tk.StringVar()
c3e1Entry = ttk.Entry(frame4, width=5, textvariable=c3e1Var)
c3e1Entry.grid(column=0, row=8, padx=2, pady=2)

c3e2Var = tk.StringVar()
c3e2Entry = ttk.Entry(frame4, width=5, textvariable=c3e2Var)
c3e2Entry.grid(column=1, row=8, padx=2, pady=2)

c3e3Var = tk.StringVar()
c3e3Entry = ttk.Entry(frame4, width=5, textvariable=c3e3Var)
c3e3Entry.grid(column=2, row=8, padx=2, pady=2)

# Label for Timing (Coupon 3)
c3TimeLabel = ttk.Label(frame4, text="Hour (24hr Format)")
c3TimeLabel.grid(column=0, row=9, columnspan=3)

# Entry for timing (Coupon 3)
c3t1Var = tk.StringVar()
c3t1Entry = ttk.Entry(frame4, width=5, textvariable=c3t1Var)
c3t1Entry.grid(column=0, row=10)

c3t2Var = tk.StringVar()
c3t2Entry = ttk.Entry(frame4, width=5, textvariable=c3t2Var)
c3t2Entry.grid(column=1, row=10, padx=2, pady=2)

c3t3Var = tk.StringVar()
c3t3Entry = ttk.Entry(frame4, width=5, textvariable=c3t3Var)
c3t3Entry.grid(column=2, row=10, padx=2, pady=2)

## Frame 5 (Buttons frame)
frame5 = ttk.Frame(win)
frame5.grid(column=0, row=2, columnspan=4, sticky="nsew")

# Button
generateButton = ttk.Button(frame5, text="Generate", command=generate)
generateButton.grid(column=0, row=2, columnspan=4, sticky="NSEW")

## Labels
helpMsg = "Fill in details and HTML for coupons and TnCs will be auto-generated. For event days, please key in a coupon description in 'Value' field (i.e., ), and paste the design in manually. Please report bugs to Gregory."
helpLabel = tk.Message(guide, text=helpMsg, width=500)
helpLabel.grid(row=0, column=0)

try:
    win.iconbitmap(resource_path("mmq.ico"))
except:
    pass

win.mainloop()
