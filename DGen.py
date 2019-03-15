import sys
import tkinter
from tkinter import ttk, messagebox, filedialog
import tkinter.font as font
from faker import Faker, Factory
import xlwt
from faker.providers import address, bank
import os
import time
import random
import json
#import pyttsx3
import getpass
import string
import csv

AU_DataGen = Faker("en_US")

AU_DataGen = Factory.create()
AU_DataGen.add_provider(address)
AU_DataGen.add_provider(bank)
student_data = {}

wb = xlwt.Workbook()
ws = wb.add_sheet("DataGenerator", cell_overwrite_ok=True)
# ws1 = wb.add_sheet("JsonData", cell_overwrite_ok=True)
# ws2 = wb.add_sheet("DBData", cell_overwrite_ok=True)
# ws1 = wb.add_sheet("JsonData", cell_overwrite_ok=True)
# ws2 = wb.add_sheet("DBData", cell_overwrite_ok=True)


# Speech Code
# engine = pyttsx3.init()
# engine.say("Hi {}, Welcome to Data Generator Tool".format(getpass.getuser()))
# engine.setProperty('rate', 50)  #120 words per minute
# engine.setProperty('volume', 0.9)
# engine.runAndWait()


# print(getpass.getuser())
# print(AU_DataGen.random_number(digits=5, fix_len=True))


root = tkinter.Tk()
root.title("Data Generator")
root.geometry("650x300+100+200")

# root.iconbitmap(r'C:\Users\rhebbar\PycharmProjects\APITest\DGIcon1.ico')

frame = tkinter.Frame(root)
frame.pack(fill="both")
logo = tkinter.PhotoImage(file="DLogo.gif")
MyFont = font.Font(size=12, family='Arial')
MyFont2 = font.Font(size=11, family='Arial')
root.resizable(False, False)

# EmailList = ["@abc.com", "@xyz.in", "@outlook1.in"]

# RandomMail = str(random.choice(EmailList))


def create_window():

    window = tkinter.Toplevel(root)
    # window.geometry("600x400+300+200")
    window.resizable(False, False)
    #window.iconbitmap(r'C:\Users\rhebbar\PycharmProjects\APITest\DGIcon1.ico')
    window.focus_set()

    topFrame = tkinter.Frame(window, highlightthickness=4, bd=4, highlightbackground="DARKGRAY", highlightcolor="DARKGRAY")
    topFrame.pack(side=tkinter.TOP, expand=1, anchor='n', pady=6, padx=10)
    titleLabel = tkinter.Label(topFrame, font=('Verdana', 12, 'bold'),
                               text="Data Generator Tool v1.0",
                               bd=5, anchor='w', padx=15)
    titleLabel.pack(side=tkinter.LEFT)

    DescInfo = "*CH - Column Header\n *CN - Column Number"

    Desc = tkinter.Label(topFrame, font=('Verdana', 8, 'bold'),
                         text=DescInfo,
                         bd=5, anchor='w', padx=30)
    Desc.pack(side=tkinter.LEFT)

    CustStrInfo = "Custom String Inputs\n# - for numbers\n ? - for alphabets"

    Desc = tkinter.Label(topFrame, font=('Verdana', 8, 'bold'),
                         text=CustStrInfo,
                         bd=5, anchor='w', padx=30, relief=tkinter.RIDGE)
    Desc.pack(side=tkinter.LEFT)

    clockFrame = tkinter.Frame(topFrame, width=100, height=50, bd=4, relief="ridge", bg='grey')
    clockFrame.pack(side=tkinter.RIGHT, padx=14, pady=8)
    clockLabel = tkinter.Label(clockFrame, font=('arial', 12, 'bold'), bd=5, anchor='e')
    clockLabel.pack()

    def tick(curtime=''):  # acts as a clock, changing the label when the time goes up
        newtime = time.strftime('%I:%M:%S %p')  # %H - for 24 hr format, %p for AM or PM
        if newtime != curtime:
            curtime = newtime
            clockLabel.config(text=curtime)
        clockLabel.after(200, tick, curtime)

    tick()

# ==========================================##############################=========================================

    mainframe1 = tkinter.Frame(window, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1 = tkinter.Frame(mainframe1, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1 = tkinter.Frame(frame1, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1a = tkinter.Frame(mainframe1, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1a.pack(side=tkinter.LEFT, fill=tkinter.BOTH, pady=2)

    subframe1a = tkinter.Frame(frame1a, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1c = tkinter.Frame(mainframe1, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1c.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1c = tkinter.Frame(frame1c, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    mainframe1.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=4)

    mainframe2 = tkinter.Frame(window, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1b = tkinter.Frame(mainframe2, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1b.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1b = tkinter.Frame(frame1b, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1d = tkinter.Frame(mainframe2, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1d.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1d = tkinter.Frame(frame1d, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    mainframe2.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=4)

    mainframe3 = tkinter.Frame(window, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1f = tkinter.Frame(mainframe3, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1f.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1f = tkinter.Frame(frame1f, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1o = tkinter.Frame(mainframe3, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1o.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1o = tkinter.Frame(frame1o, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    mainframe3.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=4)

    mainframe3a = tkinter.Frame(window, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1g = tkinter.Frame(mainframe3a, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1g.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1g = tkinter.Frame(frame1g, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1h = tkinter.Frame(mainframe3a, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1h.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=2, pady=2)

    subframe1h = tkinter.Frame(frame1h, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1i = tkinter.Frame(mainframe3a, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1i.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1i = tkinter.Frame(frame1i, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    mainframe3a.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=4)

    mainframe3b = tkinter.Frame(window, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1e = tkinter.Frame(mainframe3b, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1e.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1e = tkinter.Frame(frame1e, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1n = tkinter.Frame(mainframe3b, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1n.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=2, pady=2)

    subframe1n = tkinter.Frame(frame1n, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    mainframe3b.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=4)

    frame1n.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=2, pady=2)

    mainframe4 = tkinter.Frame(window, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1j = tkinter.Frame(mainframe4, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1j.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1j = tkinter.Frame(frame1j, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1k = tkinter.Frame(mainframe4, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1k.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1k = tkinter.Frame(frame1k, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    mainframe4.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=4)

    mainframe5 = tkinter.Frame(window, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1l = tkinter.Frame(mainframe5, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1l.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1l = tkinter.Frame(frame1l, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    frame1m = tkinter.Frame(mainframe5, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE)
    frame1m.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)

    subframe1m = tkinter.Frame(frame1m, highlightthickness=1, bd=4, pady=3, relief=tkinter.RIDGE, bg='grey')

    mainframe5.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=4)

    frame2 = tkinter.Frame(window, highlightthickness=4, bd=4, pady=3, highlightbackground="DARKGRAY", highlightcolor="DARKGRAY")
    frame2.pack(side=tkinter.TOP, fill=tkinter.BOTH, padx=10, pady=5, ipady=2)

    bottomframe = tkinter.Frame(window)
    bottomframe.pack(side=tkinter.TOP, padx=10)

    def toFrom():

        if NumRange.cget("text") == "Number_Range":
            if rnumberVar.get() == 1:
                FromLabel.pack(side="left", anchor="w", padx=5)
                FromLen.pack(side="left")
                ToLabel.pack(side="left", anchor="w", padx=5)
                ToLen.pack(side="left", padx=5)
            else:
                FromLen.pack_forget()
                ToLen.pack_forget()

        if NumLen.cget("text") == "Number_Length":
            if lnumberVar.get() == 1:
                NumLenLabel.pack(side="left", padx=5)
                rToLen.pack(side="left", padx=5)
            else:
                rToLen.pack_forget()

        if bothify.cget("text") == "Custom_String":
            if bothifyVar.get() == 1:
                CustomStr.pack(side="left", padx=5)
            else:
                CustomStr.pack_forget()

        if fullname.cget("text") == "Full_Name":
            if nameVar.get() == 1:
                subframe1.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=5, pady=2)
                CHLabel.pack(side="left", anchor="w", padx=5)
                ColumnHeader1.pack(side="left", padx=5)
                CNLabel.pack(side="left", anchor="w", padx=5)
                ColumnNumber1.pack(side="left", padx=5)
            else:
                subframe1.pack_forget()
                ColumnHeader1.pack_forget()
                ColumnNumber1.pack_forget()
                CHLabel.pack_forget()
                CNLabel.pack_forget()

        if fname.cget("text") == "First_Name":
            if fnameVar.get() == 1:
                subframe1a.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabela.pack(side="left", anchor="w", padx=5)
                fnameColumnHeader.pack(side="left")
                CNLabela.pack(side="left", anchor="w", padx=5)
                fnameColumnNumber.pack(side="left", padx=5)
            else:
                subframe1a.pack_forget()
                fnameColumnHeader.pack_forget()
                fnameColumnNumber.pack_forget()
                CHLabela.pack_forget()
                CNLabela.pack_forget()

        if lname.cget("text") == "Last_Name":
            if lnameVar.get() == 1:
                subframe1c.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelc.pack(side="left", anchor="w", padx=5)
                lnameColumnHeader.pack(side="left", padx=5)
                CNLabelc.pack(side="left", anchor="w", padx=5)
                lnameColumnNumber.pack(side="left", padx=5)
            else:
                subframe1c.pack_forget()
                lnameColumnHeader.pack_forget()
                lnameColumnNumber.pack_forget()
                CHLabelc.pack_forget()
                CNLabelc.pack_forget()

        if NumRange.cget("text") == "Number_Range":
            if rnumberVar.get() == 1:
                subframe1b.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelb.pack(side="left", anchor="w", padx=5)
                nrColumnHeader.pack(side="left", padx=5)
                CNLabelb.pack(side="left", anchor="w", padx=5)
                nrColumnNumber.pack(side="left", padx=5)
            else:
                subframe1b.pack_forget()
                nrColumnHeader.pack_forget()
                nrColumnNumber.pack_forget()
                CHLabelb.pack_forget()
                CNLabelb.pack_forget()

        if NumLen.cget("text") == "Number_Length":
            if lnumberVar.get() == 1:
                subframe1d.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabeld.pack(side="left", anchor="w", padx=5)
                nlColumnHeader.pack(side="left", padx=5)
                CNLabeld.pack(side="left", anchor="w", padx=5)
                nlColumnNumber.pack(side="left", padx=5)
            else:
                subframe1d.pack_forget()
                nlColumnHeader.pack_forget()
                nlColumnNumber.pack_forget()
                CHLabeld.pack_forget()
                CNLabeld.pack_forget()

        if email.cget("text") == "Email":
            if emailVar.get() == 1:
                subframe1e.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabele.pack(side="left", anchor="w", padx=5)
                emailColumnHeader.pack(side="left", padx=5)
                CNLabele.pack(side="left", anchor="w", padx=5)
                emailColumnNumber.pack(side="left", padx=5)
            else:
                subframe1e.pack_forget()
                emailColumnHeader.pack_forget()
                emailColumnNumber.pack_forget()
                CHLabele.pack_forget()
                CNLabele.pack_forget()

        if address.cget("text") == "Address Line 1":
            if addressVar.get() == 1:
                subframe1f.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelf.pack(side="left", anchor="w", padx=5)
                addressColumnHeader.pack(side="left", padx=5)
                CNLabelf.pack(side="left", anchor="w", padx=5)
                addressColumnNumber.pack(side="left", padx=5)
            else:
                subframe1f.pack_forget()
                addressColumnHeader.pack_forget()
                addressColumnNumber.pack_forget()
                CHLabelf.pack_forget()
                CNLabelf.pack_forget()

        if country.cget("text") == "Country":
            if countryVar.get() == 1:
                subframe1g.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelg.pack(side="left", anchor="w", padx=5)
                countryColumnHeader.pack(side="left", padx=5)
                CNLabelg.pack(side="left", anchor="w", padx=5)
                countryColumnNumber.pack(side="left", padx=5)
            else:
                subframe1g.pack_forget()
                countryColumnHeader.pack_forget()
                countryColumnNumber.pack_forget()
                CHLabelg.pack_forget()
                CNLabelg.pack_forget()

        if city.cget("text") == "City":
            if cityVar.get() == 1:
                subframe1h.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelh.pack(side="left", anchor="w", padx=5)
                cityColumnHeader.pack(side="left", padx=5)
                CNLabelh.pack(side="left", anchor="w", padx=5)
                cityColumnNumber.pack(side="left", padx=5)
            else:
                subframe1h.pack_forget()
                cityColumnHeader.pack_forget()
                cityColumnNumber.pack_forget()
                CHLabelh.pack_forget()
                CNLabelh.pack_forget()

        if zipcode.cget("text") == "Zipcode":
            if zipCodeVar.get() == 1:
                subframe1i.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabeli.pack(side="left", anchor="w", padx=5)
                zipcodeColumnHeader.pack(side="left", padx=5)
                CNLabeli.pack(side="left", anchor="w", padx=5)
                zipcodeColumnNumber.pack(side="left", padx=5)
            else:
                subframe1i.pack_forget()
                zipcodeColumnHeader.pack_forget()
                zipcodeColumnNumber.pack_forget()
                CHLabeli.pack_forget()
                CNLabeli.pack_forget()

        if bothify.cget("text") == "Custom_String":
            if bothifyVar.get() == 1:
                subframe1j.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelj.pack(side="left", anchor="w", padx=5)
                bothifyColumnHeader.pack(side="left", padx=5)
                CNLabelj.pack(side="left", anchor="w", padx=5)
                bothifyColumnNumber.pack(side="left", padx=5)
            else:
                subframe1j.pack_forget()
                bothifyColumnHeader.pack_forget()
                bothifyColumnNumber.pack_forget()
                CHLabelj.pack_forget()
                CNLabelj.pack_forget()

        if decimal.cget("text") == "Decimal Number":
            if decimalVar.get() == 1:
                subframe1k.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                LeftDecimalLabel.pack(side="left", anchor="w", padx=5)
                LeftDecimal.pack(side="left", padx=5)
                RightDecimalLabel.pack(side="left", anchor="w", padx=5)
                RightDecimal.pack(side="left", padx=5)
                CHLabelk.pack(side="left", anchor="w", padx=5)
                decimalColumnHeader.pack(side="left", padx=5)
                CNLabelk.pack(side="left", anchor="w", padx=5)
                decimalColumnNumber.pack(side="left", padx=5)
            else:
                subframe1k.pack_forget()
                decimalColumnHeader.pack_forget()
                decimalColumnNumber.pack_forget()
                CHLabelk.pack_forget()
                CNLabelk.pack_forget()

        if datelabel.cget("text") == "Date":
            if dateVar.get() == 1:
                subframe1l.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabell.pack(side="left", anchor="w", padx=5)
                dateColumnHeader.pack(side="left", padx=5)
                CNLabell.pack(side="left", anchor="w", padx=5)
                dateColumnNumber.pack(side="left", padx=5)
            else:
                subframe1l.pack_forget()
                dateColumnHeader.pack_forget()
                dateColumnNumber.pack_forget()
                CHLabell.pack_forget()
                CNLabell.pack_forget()

        if timelabel.cget("text") == "Time":
            if timevar.get() == 1:
                subframe1m.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelm.pack(side="left", anchor="w", padx=5)
                timeColumnHeader.pack(side="left", padx=5)
                CNLabelm.pack(side="left", anchor="w", padx=5)
                timeColumnNumber.pack(side="left", padx=5)
            else:
                subframe1m.pack_forget()
                timeColumnHeader.pack_forget()
                timeColumnNumber.pack_forget()
                CHLabelm.pack_forget()
                CNLabelm.pack_forget()

        if phone.cget("text") == "Phone Number":
            if phoneVar.get() == 1:
                subframe1n.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabeln.pack(side="left", anchor="w", padx=5)
                phoneColumnHeader.pack(side="left", padx=5)
                CNLabeln.pack(side="left", anchor="w", padx=5)
                phoneColumnNumber.pack(side="left", padx=5)
            else:
                subframe1n.pack_forget()
                phoneColumnHeader.pack_forget()
                phoneColumnNumber.pack_forget()
                CHLabeln.pack_forget()
                CNLabeln.pack_forget()

        if addressline2.cget("text") == "Address Line 2":
            if addressLine2Var.get() == 1:
                subframe1o.pack(side=tkinter.LEFT, fill=tkinter.BOTH, padx=10, pady=2)
                CHLabelo.pack(side="left", anchor="w", padx=5)
                addrLine2ColumnHeader.pack(side="left", padx=5)
                CNLabelo.pack(side="left", anchor="w", padx=5)
                addrLine2ColumnNumber.pack(side="left", padx=5)
            else:
                subframe1o.pack_forget()
                addrLine2ColumnHeader.pack_forget()
                addrLine2ColumnNumber.pack_forget()
                CHLabelo.pack_forget()
                CNLabelo.pack_forget()

    def toCombo(Self):
        if dateComboVar.get() == "Other Formats":
            OtherFormarLabel.pack(side="left", padx=5)
            OtherDateFormat.pack(side="left", padx=5)
        else:
            OtherFormarLabel.pack_forget()
            OtherDateFormat.pack_forget()

    # ===================================#############################+===========================================

    def callback():
        if len(ranger.get()) == 0 or not str.isnumeric(ranger.get()):
            messagebox.showerror("Validation Error", "Please enter loop value(integer)")
        elif int(ranger.get()) <= 0:
            messagebox.showerror("Validation Error", "Please enter loop value(integer) > 0")
        else:
            pass

        # for x in range(int(ranger.get())):
            if fullname.cget("text") == "Full_Name":
                if nameVar.get() == 1:
                    # global count
                    hasPrinted = False
                    if len(ColumnHeader1.get()) == 0 or len(ColumnNumber1.get()) == 0:
                        messagebox.showerror("Validation Error", "'Full Name' entrybox should not be empty")
                    elif not str.isnumeric(ColumnNumber1.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Full Name' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.name())
                            ws.write(0, int(ColumnNumber1.get()), ColumnHeader1.get())
                            ws.write(x + 1, int(ColumnNumber1.get()), AU_DataGen.name())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(fullname.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if fname.cget("text") == "First_Name":
                if fnameVar.get() == 1:
                    hasPrinted = False
                    if len(fnameColumnHeader.get()) == 0 or len(fnameColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'First Name' entrybox should not be empty")
                    elif not str.isnumeric(fnameColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'First Name' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.first_name())
                            ws.write(0, int(fnameColumnNumber.get()), fnameColumnHeader.get())
                            ws.write(x + 1, int(fnameColumnNumber.get()), AU_DataGen.first_name())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(fname.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass
                    # print("input please")

            if lname.cget("text") == "Last_Name":
                if lnameVar.get() == 1:
                    hasPrinted = False
                    if len(lnameColumnHeader.get()) == 0 or len(lnameColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error", "'Last Name' entrybox should not be empty")
                    elif not str.isnumeric(lnameColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Last Name' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.last_name())
                            ws.write(0, int(lnameColumnNumber.get()), lnameColumnHeader.get())
                            ws.write(x + 1, int(lnameColumnNumber.get()), AU_DataGen.last_name())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(lname.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if NumRange.cget("text") == "Number_Range":
                if rnumberVar.get() == 1:
                    hasPrinted = False
                    if len(nrColumnHeader.get()) == 0 or len(nrColumnNumber.get()) == 0 or len(FromLen.get()) == 0 or len(
                            ToLen.get()) == 0:
                        messagebox.showerror("Validation Error", "'Number Range' entrybox should not be empty")
                    elif not str.isnumeric(nrColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Number Range' CN textbox")
                    elif not str.isnumeric(FromLen.get()) or not str.isnumeric(ToLen.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Number Length' From and To textbox")
                    elif ToLen.get() <= FromLen.get():
                        messagebox.showerror("Validation Error",
                                            "'From' value should not be greater than or equal to 'To' Value")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.random.randint(int(FromLen.get()), int(ToLen.get())))
                            ws.write(0, int(nrColumnNumber.get()), nrColumnHeader.get())
                            ws.write(x + 1, int(nrColumnNumber.get()),
                                     AU_DataGen.random.randint(int(FromLen.get()), int(ToLen.get())))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(NumRange.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if NumLen.cget("text") == "Number_Length":
                if lnumberVar.get() == 1:
                    hasPrinted = False
                    if len(nlColumnHeader.get()) == 0 or len(nlColumnNumber.get()) == 0 or len(rToLen.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Number Length' entrybox should not be empty")
                    elif not str.isnumeric(nlColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Number Length' CN textbox")
                    elif not str.isnumeric(rToLen.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Number Length' Length textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.random_number(digits=int(rToLen.get()), fix_len=True))
                            ws.write(0, int(nlColumnNumber.get()), nlColumnHeader.get())
                            ws.write(x + 1, int(nlColumnNumber.get()),
                                     AU_DataGen.random_number(digits=int(rToLen.get()), fix_len=True))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(NumLen.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if email.cget("text") == "Email":
                if emailVar.get() == 1:
                    hasPrinted = False
                    if len(emailColumnHeader.get()) == 0 or len(emailColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Email' entrybox should not be empty")
                    elif not str.isnumeric(emailColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Email' CN textbox")

                    elif safe_emailVar.get() == 1:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.safe_email())
                            ws.write(0, int(emailColumnNumber.get()), emailColumnHeader.get())
                            ws.write(x + 1, int(emailColumnNumber.get()), AU_DataGen.safe_email())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(email.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.email())
                            ws.write(0, int(emailColumnNumber.get()), emailColumnHeader.get())
                            ws.write(x + 1, int(emailColumnNumber.get()), AU_DataGen.email())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(email.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if address.cget("text") == "Address Line 1":
                if addressVar.get() == 1:
                    hasPrinted = False
                    if len(addressColumnHeader.get()) == 0 or len(addressColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Address' entrybox should not be empty")
                    elif not str.isnumeric(addressColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Address Line 1' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.address())
                            ws.write(0, int(addressColumnNumber.get()), addressColumnHeader.get())
                            ws.write(x + 1, int(addressColumnNumber.get()), AU_DataGen.address())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(address.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if country.cget("text") == "Country":
                if countryVar.get() == 1:
                    hasPrinted = False
                    if len(countryColumnHeader.get()) == 0 or len(countryColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Country' entrybox should not be empty")
                    elif not str.isnumeric(countryColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Country' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.country())
                            ws.write(0, int(countryColumnNumber.get()), countryColumnHeader.get())
                            ws.write(x + 1, int(countryColumnNumber.get()), AU_DataGen.country())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(country.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if city.cget("text") == "City":
                if cityVar.get() == 1:
                    hasPrinted = False
                    if len(cityColumnHeader.get()) == 0 or len(cityColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'City' entrybox should not be empty")
                    elif not str.isnumeric(cityColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'City' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.city())
                            ws.write(0, int(cityColumnNumber.get()), cityColumnHeader.get())
                            ws.write(x + 1, int(cityColumnNumber.get()), AU_DataGen.city())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(city.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if zipcode.cget("text") == "Zipcode":
                if zipCodeVar.get() == 1:
                    hasPrinted = False
                    if len(zipcodeColumnHeader.get()) == 0 or len(zipcodeColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Zipcode' entrybox should not be empty")
                    elif not str.isnumeric(zipcodeColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Zipcode' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.zipcode())
                            ws.write(0, int(zipcodeColumnNumber.get()), zipcodeColumnHeader.get())
                            ws.write(x + 1, int(zipcodeColumnNumber.get()), AU_DataGen.zipcode())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(zipcode.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if bothify.cget("text") == "Custom_String":
                if bothifyVar.get() == 1:
                    hasPrinted = False
                    if len(bothifyColumnHeader.get()) == 0 or len(bothifyColumnNumber.get()) == 0 or len(CustomStr.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Custom String' entrybox should not be empty")
                    elif not str.isnumeric(bothifyColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Custom_String' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.bothify(text=CustomStr.get()))
                            ws.write(0, int(bothifyColumnNumber.get()), bothifyColumnHeader.get())
                            ws.write(x + 1, int(bothifyColumnNumber.get()), AU_DataGen.bothify(text=CustomStr.get()))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(bothify.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if decimal.cget("text") == "Decimal Number":
                if decimalVar.get() == 1:
                    hasPrinted = False
                    if len(decimalColumnHeader.get()) == 0 or len(decimalColumnNumber.get()) == 0 or len(LeftDecimal.get()) == 0 or len(RightDecimal.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Custom String' entrybox should not be empty")
                    elif not str.isnumeric(decimalColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Decimal Number' CN textbox")
                    elif not str.isnumeric(LeftDecimal.get()) or not str.isnumeric(RightDecimal.get()):
                        messagebox.showerror("Validation Error",
                                             "Please enter integer value in 'Decimal Number' Left and Right Decimal textbox")
                    elif int(LeftDecimal.get()) < 0:
                        messagebox.showerror("Validation Error",
                                             "Left Decimal value should be greater than 0")
                    elif int(LeftDecimal.get()) == 0 and int(RightDecimal.get()) == 0:
                        messagebox.showerror("Validation Error",
                                             "Both Left Decimal and Right Decimal value shoild not be equal to 0")
                    elif decimaPositivelVar.get() == 1:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.pydecimal(left_digits=int(LeftDecimal.get()), right_digits=int(RightDecimal.get()), positive=True))
                            ws.write(0, int(decimalColumnNumber.get()), decimalColumnHeader.get())
                            ws.write(x + 1, int(decimalColumnNumber.get()), AU_DataGen.pydecimal(left_digits=int(LeftDecimal.get()), right_digits=int(RightDecimal.get()), positive=True))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(decimal.cget("text")))
                            hasPrinted = True
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.pydecimal(left_digits=int(LeftDecimal.get()), right_digits=int(RightDecimal.get()), positive=False))
                            ws.write(0, int(decimalColumnNumber.get()), decimalColumnHeader.get())
                            ws.write(x + 1, int(decimalColumnNumber.get()), AU_DataGen.pydecimal(left_digits=int(LeftDecimal.get()), right_digits=int(RightDecimal.get()), positive=False))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(bothify.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if timelabel.cget("text") == "Time":
                if timevar.get() == 1:
                    hasPrinted = False
                    if len(timeColumnHeader.get()) == 0 or len(timeColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Time' entrybox should not be empty")
                    elif not str.isnumeric(timeColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Time' CN textbox")
                    elif timeFormatVar.get() == 1:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.time(pattern="%I:%M:%S %p", end_datetime=None))
                            ws.write(0, int(timeColumnNumber.get()), timeColumnHeader.get())
                            ws.write(x + 1, int(timeColumnNumber.get()), AU_DataGen.time(pattern="%I:%M:%S %p", end_datetime=None))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(timelabel.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.time(pattern="%H:%M:%S", end_datetime=None))
                            ws.write(0, int(timeColumnNumber.get()), timeColumnHeader.get())
                            ws.write(x + 1, int(timeColumnNumber.get()), AU_DataGen.time(pattern="%H:%M:%S", end_datetime=None))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(timelabel.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if datelabel.cget("text") == "Date":
                if dateVar.get() == 1:
                    hasPrinted = False
                    if len(dateColumnHeader.get()) == 0 or len(dateColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Date' entrybox should not be empty")
                    elif not str.isnumeric(dateColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Date' CN textbox")

                    elif dateComboVar.get() == "Select Format":
                        messagebox.showerror("Validation Error", "Please select format from dropdown in 'Date'")

                    elif dateComboVar.get() == "dd-mm-yyyy":
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.date(pattern="%d-%m-%Y", end_datetime=None))
                            ws.write(0, int(dateColumnNumber.get()), dateColumnHeader.get())
                            ws.write(x + 1, int(dateColumnNumber.get()), AU_DataGen.date(pattern="%d-%m-%Y", end_datetime=None))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(datelabel.cget("text")))
                            hasPrinted = True
                        else:
                            pass

                    elif dateComboVar.get() == "mm-dd-yyyy":
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.date(pattern="%m-%d-%Y", end_datetime=None))
                            ws.write(0, int(dateColumnNumber.get()), dateColumnHeader.get())
                            ws.write(x + 1, int(dateColumnNumber.get()), AU_DataGen.date(pattern="%m-%d-%Y", end_datetime=None))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(datelabel.cget("text")))
                            hasPrinted = True
                        else:
                            pass

                    elif dateComboVar.get() == "dd-(mon)-yyyy":
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.date(pattern="%d-%b-%Y", end_datetime=None))
                            ws.write(0, int(dateColumnNumber.get()), dateColumnHeader.get())
                            ws.write(x + 1, int(dateColumnNumber.get()), AU_DataGen.date(pattern="%d-%b-%Y", end_datetime=None))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(datelabel.cget("text")))
                            hasPrinted = True
                        else:
                            pass

                    elif dateComboVar.get() == "dd-(month)-yyyy":
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.date(pattern="%d-%B-%Y", end_datetime=None))
                            ws.write(0, int(dateColumnNumber.get()), dateColumnHeader.get())
                            ws.write(x + 1, int(dateColumnNumber.get()), AU_DataGen.date(pattern="%d-%B-%Y", end_datetime=None))
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(datelabel.cget("text")))
                            hasPrinted = True
                        else:
                            pass

                    elif dateComboVar.get() == "Other Formats":
                        if len(OtherDateFormat.get()) == 0:
                            messagebox.showerror("Validation Error", "'Custom Fromat' entrybox should not be empty")
                        else:
                            for x in range(int(ranger.get())):
                                print(AU_DataGen.date(pattern="{}".format(OtherDateFormat.get()), end_datetime=None))
                                ws.write(0, int(dateColumnNumber.get()), dateColumnHeader.get())
                                ws.write(x + 1, int(dateColumnNumber.get()), AU_DataGen.date(pattern="{}".format(OtherDateFormat.get()), end_datetime=None))
                            if not hasPrinted:
                                messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(datelabel.cget("text")))
                                hasPrinted = True
                            else:
                                pass

                    else:
                        pass
                else:
                    pass

            if phone.cget("text") == "Phone Number":
                if phoneVar.get() == 1:
                    hasPrinted = False
                    if len(phoneColumnHeader.get()) == 0 or len(phoneColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Phone Number' entrybox should not be empty")
                    elif not str.isnumeric(phoneColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Phone Number' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.phone_number())
                            ws.write(0, int(phoneColumnNumber.get()), phoneColumnHeader.get())
                            ws.write(x + 1, int(phoneColumnNumber.get()), AU_DataGen.phone_number())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(phone.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

            if addressline2.cget("text") == "Address Line 2":
                if addressLine2Var.get() == 1:
                    hasPrinted = False
                    if len(addrLine2ColumnHeader.get()) == 0 or len(addrLine2ColumnNumber.get()) == 0:
                        messagebox.showerror("Validation Error",
                                            "'Address Line 2' entrybox should not be empty")
                    elif not str.isnumeric(addrLine2ColumnNumber.get()):
                        messagebox.showerror("Validation Error", "Please enter integer value in 'Address Line 2' CN textbox")
                    else:
                        for x in range(int(ranger.get())):
                            print(AU_DataGen.address())
                            ws.write(0, int(addrLine2ColumnNumber.get()), addrLine2ColumnHeader.get())
                            ws.write(x + 1, int(addrLine2ColumnNumber.get()), AU_DataGen.address())
                        if not hasPrinted:
                            messagebox.showinfo('sucess info', 'Data Generated Successfully for {}'.format(addressline2.cget("text")))
                            hasPrinted = True
                        else:
                            pass
                else:
                    pass

        window.focus_set()
    # =====================================##########################################+===========================

    def browsefunc():
        dir_path = os.path.expanduser('~/Documents')
        filename = filedialog.asksaveasfilename(initialdir="{}".format(dir_path), defaultextension=".xls")
        wb.save(filename)

    browsebutton = tkinter.Button(frame2, text="Save File", relief=tkinter.GROOVE, command=browsefunc)
    browsebutton.pack(side="right", padx=10)

    # wb.save("EXCEL/AutoDataGenerator1{}.xls".format(time.strftime("%Y%m%d-%H%M%S")))

    # =====================================##########################################+===========================

    FromLen = ttk.Entry(subframe1b, width=10)

    ToLen = ttk.Entry(subframe1b, width=10)

    rToLen = ttk.Entry(subframe1d, width=4)

    CustomStr = ttk.Entry(frame1j)

    LeftDecimal = ttk.Entry(subframe1k, width=6)

    RightDecimal = ttk.Entry(subframe1k, width=6)

    OtherDateFormat = ttk.Entry(frame1l, width=16)

    ColumnHeader1 = ttk.Entry(subframe1, width=16)

    ColumnNumber1 = ttk.Entry(subframe1, width=4)

    fnameColumnHeader = ttk.Entry(subframe1a, width=16)

    fnameColumnNumber = ttk.Entry(subframe1a, width=4)

    lnameColumnHeader = ttk.Entry(subframe1c, width=16)

    lnameColumnNumber = ttk.Entry(subframe1c, width=4)

    nrColumnHeader = ttk.Entry(subframe1b, width=16)

    nrColumnNumber = ttk.Entry(subframe1b, width=4)

    nlColumnHeader = ttk.Entry(subframe1d, width=16)

    nlColumnNumber = ttk.Entry(subframe1d, width=4)

    emailColumnHeader = ttk.Entry(subframe1e, width=16)

    emailColumnNumber = ttk.Entry(subframe1e, width=4)

    addressColumnHeader = ttk.Entry(subframe1f, width=16)

    addressColumnNumber = ttk.Entry(subframe1f, width=4)

    countryColumnHeader = ttk.Entry(subframe1g, width=16)

    countryColumnNumber = ttk.Entry(subframe1g, width=4)

    cityColumnHeader = ttk.Entry(subframe1h, width=16)

    cityColumnNumber = ttk.Entry(subframe1h, width=4)

    zipcodeColumnHeader = ttk.Entry(subframe1i, width=16)

    zipcodeColumnNumber = ttk.Entry(subframe1i, width=4)

    bothifyColumnHeader = ttk.Entry(subframe1j, width=16)

    bothifyColumnNumber = ttk.Entry(subframe1j, width=4)

    decimalColumnHeader = ttk.Entry(subframe1k, width=16)

    decimalColumnNumber = ttk.Entry(subframe1k, width=4)

    dateColumnHeader = ttk.Entry(subframe1l, width=16)

    dateColumnNumber = ttk.Entry(subframe1l, width=4)

    timeColumnHeader = ttk.Entry(subframe1m, width=16)

    timeColumnNumber = ttk.Entry(subframe1m, width=4)

    phoneColumnHeader = ttk.Entry(subframe1n, width=16)

    phoneColumnNumber = ttk.Entry(subframe1n, width=4)

    addrLine2ColumnHeader = ttk.Entry(subframe1o, width=16)

    addrLine2ColumnNumber = ttk.Entry(subframe1o, width=4)

    FromLabel = tkinter.Label(subframe1b, text="From:")
    ToLabel = tkinter.Label(subframe1b, text="To:")

    NumLenLabel = tkinter.Label(subframe1d, text="Length")

    LeftDecimalLabel = tkinter.Label(subframe1k, text="Left Digits")
    RightDecimalLabel = tkinter.Label(subframe1k, text="Right Digits")

    OtherFormarLabel = tkinter.Label(frame1l, text="Custom Format:")

    CHLabel = tkinter.Label(subframe1, text="CH")
    CNLabel = tkinter.Label(subframe1, text="CN")

    CHLabela = tkinter.Label(subframe1a, text="CH")
    CNLabela = tkinter.Label(subframe1a, text="CN")

    CHLabelc = tkinter.Label(subframe1c, text="CH")
    CNLabelc = tkinter.Label(subframe1c, text="CN")

    CHLabelb = tkinter.Label(subframe1b, text="CH")
    CNLabelb = tkinter.Label(subframe1b, text="CN")

    CHLabeld = tkinter.Label(subframe1d, text="CH")
    CNLabeld = tkinter.Label(subframe1d, text="CN")

    CHLabele = tkinter.Label(subframe1e, text="CH")
    CNLabele = tkinter.Label(subframe1e, text="CN")

    CHLabelf = tkinter.Label(subframe1f, text="CH")
    CNLabelf = tkinter.Label(subframe1f, text="CN")

    CHLabelg = tkinter.Label(subframe1g, text="CH")
    CNLabelg = tkinter.Label(subframe1g, text="CN")

    CHLabelh = tkinter.Label(subframe1h, text="CH")
    CNLabelh = tkinter.Label(subframe1h, text="CN")

    CHLabeli = tkinter.Label(subframe1i, text="CH")
    CNLabeli = tkinter.Label(subframe1i, text="CN")

    CHLabelj = tkinter.Label(subframe1j, text="CH")
    CNLabelj = tkinter.Label(subframe1j, text="CN")

    CHLabelk = tkinter.Label(subframe1k, text="CH")
    CNLabelk = tkinter.Label(subframe1k, text="CN")

    CHLabell = tkinter.Label(subframe1l, text="CH")
    CNLabell = tkinter.Label(subframe1l, text="CN")

    CHLabelm = tkinter.Label(subframe1m, text="CH")
    CNLabelm = tkinter.Label(subframe1m, text="CN")

    CHLabeln = tkinter.Label(subframe1n, text="CH")
    CNLabeln = tkinter.Label(subframe1n, text="CN")

    CHLabelo = tkinter.Label(subframe1o, text="CH")
    CNLabelo = tkinter.Label(subframe1o, text="CN")

    nameVar = tkinter.IntVar()
    nchk = tkinter.Checkbutton(frame1, text='', variable=nameVar, command=toFrom)
    nchk.pack(side="left", padx=5)

    fullname = ttk.Label(frame1, text="Full_Name", font=('arial', 10))
    fullname.cget("text")
    fullname.pack(side="left", padx=5)

    fnameVar = tkinter.IntVar()
    fchk = tkinter.Checkbutton(frame1a, text='', variable=fnameVar, command=toFrom)
    fchk.pack(side="left", padx=5)

    fname = ttk.Label(frame1a, text="First_Name", font=('arial', 10))
    fname.cget("text")
    fname.pack(side="left", padx=5)

    lnameVar = tkinter.IntVar()
    lchk = tkinter.Checkbutton(frame1c, text='', variable=lnameVar, command=toFrom)
    lchk.pack(side="left", padx=5)

    lname = ttk.Label(frame1c, text="Last_Name", font=('arial', 10))
    lname.cget("text")
    lname.pack(side="left", padx=5)

    rnumberVar = tkinter.IntVar()
    rnchk = tkinter.Checkbutton(frame1b, text='', variable=rnumberVar, command=toFrom)
    rnchk.pack(side="left", padx=5)

    NumRange = ttk.Label(frame1b, text="Number_Range", font=('arial', 10))
    NumRange.cget("text")
    NumRange.pack(side="left", padx=5)

    lnumberVar = tkinter.IntVar()
    lnchk = tkinter.Checkbutton(frame1d, text='', variable=lnumberVar, command=toFrom)
    lnchk.pack(side="left", padx=5)

    NumLen = ttk.Label(frame1d, text="Number_Length", font=('arial', 10))
    NumLen.cget("text")
    NumLen.pack(side="left", padx=5)

    emailVar = tkinter.IntVar()
    emchk = tkinter.Checkbutton(frame1e, text='', variable=emailVar, command=toFrom)
    emchk.pack(side="left", padx=5)

    email = ttk.Label(frame1e, text="Email", font=('arial', 10))
    email.cget("text")
    email.pack(side="left", padx=5)

    safe_emailVar = tkinter.IntVar()
    safe_emchk = tkinter.Checkbutton(subframe1e, text='', variable=safe_emailVar, command=toFrom)
    safe_emchk.pack(side="left", padx=5)

    safe_email = ttk.Label(subframe1e, text="Safe_Email", font=('arial', 10))
    safe_email.cget("text")
    safe_email.pack(side="left", padx=5)

    addressVar = tkinter.IntVar()
    addchk = tkinter.Checkbutton(frame1f, text='', variable=addressVar, command=toFrom)
    addchk.pack(side="left", padx=5)

    address = ttk.Label(frame1f, text="Address Line 1", font=('arial', 10))
    address.cget("text")
    address.pack(side="left", padx=5)

    countryVar = tkinter.IntVar()
    countrychk = tkinter.Checkbutton(frame1g, text='', variable=countryVar, command=toFrom)
    countrychk.pack(side="left", padx=5)

    country = ttk.Label(frame1g, text="Country", font=('arial', 10))
    country.cget("text")
    country.pack(side="left", padx=5)

    cityVar = tkinter.IntVar()
    citychk = tkinter.Checkbutton(frame1h, text='', variable=cityVar, command=toFrom)
    citychk.pack(side="left", padx=5)

    city = ttk.Label(frame1h, text="City", font=('arial', 10))
    city.cget("text")
    city.pack(side="left", padx=5)

    zipCodeVar = tkinter.IntVar()
    zipcodechk = tkinter.Checkbutton(frame1i, text='', variable=zipCodeVar, command=toFrom)
    zipcodechk.pack(side="left", padx=5)

    zipcode = ttk.Label(frame1i, text="Zipcode", font=('arial', 10))
    zipcode.cget("text")
    zipcode.pack(side="left", padx=5)

    bothifyVar = tkinter.IntVar()
    bothifychk = tkinter.Checkbutton(frame1j, text='', variable=bothifyVar, command=toFrom)
    bothifychk.pack(side="left", padx=5)

    bothify = ttk.Label(frame1j, text="Custom_String", font=('arial', 10))
    bothify.cget("text")
    bothify.pack(side="left", padx=5)

    decimalVar = tkinter.IntVar()
    decimalchk = tkinter.Checkbutton(frame1k, text='', variable=decimalVar, command=toFrom)
    decimalchk.pack(side="left", padx=5)

    decimal = ttk.Label(frame1k, text="Decimal Number", font=('arial', 10))
    decimal.cget("text")
    decimal.pack(side="left", padx=5)

    decimaPositivelVar = tkinter.IntVar()
    decimalPositivechk = tkinter.Checkbutton(subframe1k, text='Only Positive', variable=decimaPositivelVar, command=toFrom)
    decimalPositivechk.pack(side="left", padx=5)

    dateVar = tkinter.IntVar()
    datechk = tkinter.Checkbutton(frame1l, text='', variable=dateVar, command=toFrom)
    datechk.pack(side="left", padx=5)

    datelabel = ttk.Label(frame1l, text="Date", font=('arial', 10))
    datelabel.cget("text")
    datelabel.pack(side="left", padx=5)

    timevar = tkinter.IntVar()
    timechk = tkinter.Checkbutton(frame1m, text='', variable=timevar, command=toFrom)
    timechk.pack(side="left", padx=5)

    timelabel = ttk.Label(frame1m, text="Time", font=('arial', 10))
    timelabel.cget("text")
    timelabel.pack(side="left", padx=5)

    timeFormatVar = tkinter.IntVar()
    timeformatchk = tkinter.Checkbutton(subframe1m, text='12 Hour Format', variable=timeFormatVar, command=toFrom)
    timeformatchk.pack(side="left", padx=5)

    dateComboVar = tkinter.StringVar(frame1l)
    dateComboVar.set("Select Format")

    DateCombo = tkinter.OptionMenu(frame1l, dateComboVar,  "dd-mm-yyyy",  "mm-dd-yyyy", "dd-(mon)-yyyy", "dd-(month)-yyyy", "Other Formats", command=toCombo)
    DateCombo.pack(side="left", padx=5)

    phoneVar = tkinter.IntVar()
    phonechk = tkinter.Checkbutton(frame1n, text='', variable=phoneVar, command=toFrom)
    phonechk.pack(side="left", padx=5)

    phone = ttk.Label(frame1n, text="Phone Number", font=('arial', 10))
    phone.cget("text")
    phone.pack(side="left", padx=5)

    addressLine2Var = tkinter.IntVar()
    addressline2chk = tkinter.Checkbutton(frame1o, text='', variable=addressLine2Var, command=toFrom)
    addressline2chk.pack(side="left", padx=5)

    addressline2 = ttk.Label(frame1o, text="Address Line 2", font=('arial', 10))
    addressline2.cget("text")
    addressline2.pack(side="left", padx=5)

    # window.focus_set()
    # ==============================================#######################=========================

    fp = os.path.abspath(os.path.expanduser('~/Documents'))

    def Filepath(Self):
        os.startfile(fp)

    # =============================================#####################################+===========================

    RangeLabel = tkinter.Label(frame2, text="Number of Times")
    RangeLabel.pack(side="left", anchor="w", padx=10)

    ranger = ttk.Entry(frame2)
    ranger.pack(side="left")
    ranger.config(width=10)

    link = tkinter.Label(frame2, text="Excel Folder", cursor="hand2", fg="blue")
    link.bind("<Button-1>", Filepath)
    link.pack(side="right", padx=15)

    b = tkinter.Button(frame2, text="Generate", bg="GAINSBORO", width=10, borderwidth="3px", command=callback)
    b.pack(side="right", padx=5, )


w1 = tkinter.Label(root, image=logo, relief=tkinter.RAISED, borderwidth="5px").pack(side="right", fill="both", expand=1)

explanation = """Generate data for various data types in one go.\n\n\n
© 2018 QACC Tools Team Rights Reserved."""

w2 = tkinter.Label(root, justify=tkinter.LEFT, relief=tkinter.RIDGE, padx=15, pady=15, text=explanation,
                   font=MyFont2, borderwidth="5px").pack(side="left", fill="both")

button = tkinter.Button(frame,
                        text="Welcome to Data Generator Tool",
                        fg="black",
                        command=create_window,
                        height=2, width=30,
                        bg="GAINSBORO",
                        borderwidth="4px"
                        )
button.config(relief=tkinter.RAISED)
# button.grid(row=0, column=0, sticky=tkinter.W+tkinter.E)
button.pack(side=tkinter.BOTTOM, fill="both")
button['font'] = MyFont

root.mainloop()


