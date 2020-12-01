from tkinter import *
from tkinter import messagebox, ttk, scrolledtext
from decimal import ROUND_HALF_UP, Decimal
from win32com import client
from datetime import datetime
from pathlib import Path
import subprocess as cmd
import sqlite3
import re
import xlsxwriter
import win32api
import os


charcoal = '#F1F1EE'
rust = '#F62A00'
navy = '#00293C'
teal = '#1E656D'
clients = []
purchasers = []
products = []
years = ["2020-2021", "2021-2022", "2022-2023",  "2023-2024",
         "2024-2025", "2025-2026", "2026-2027", "2027-2028", "2028-2029", "2029-2030"]
months = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
          7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'}
one = ["", "One ", "Two ", "Three ", "Four ", "Five ", "Six ", "Seven ", "Eight ", "Nine ", "Ten ", "Eleven ",
       "Twelve ", "Thirteen ", "Fourteen ", "Fifteen ", "Sixteen ", "Seventeen ", "Eighteen ", "Nineteen "]
ten = ["", "", "Twenty ", "Thirty ", "Fourty ",
       "Fifty ", "Sixty ", "Seventy ", "Eighty ", "Ninety "]


# XLSXWriter
def draw_frame_border(workbook, worksheet, first_row, first_col, rows_count, cols_count, thickness=1):

    if cols_count == 1 and rows_count == 1:
        # whole cell
        worksheet.conditional_format(first_row, first_col,
                                     first_row, first_col,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'top': thickness, 'bottom': thickness,
                                                                     'left': thickness, 'right': thickness})})
    elif rows_count == 1:
        # left cap
        worksheet.conditional_format(first_row, first_col,
                                     first_row, first_col,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'top': thickness, 'left': thickness, 'bottom': thickness})})
        # top and bottom sides
        worksheet.conditional_format(first_row, first_col + 1,
                                     first_row, first_col + cols_count - 2,
                                     {'type': 'formula', 'criteria': 'True', 'format': workbook.add_format({'top': thickness, 'bottom': thickness})})

        # right cap
        worksheet.conditional_format(first_row, first_col + cols_count - 1,
                                     first_row, first_col + cols_count - 1,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'top': thickness, 'right': thickness, 'bottom': thickness})})

    elif cols_count == 1:
        # top cap
        worksheet.conditional_format(first_row, first_col,
                                     first_row, first_col,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'top': thickness, 'left': thickness, 'right': thickness})})

        # left and right sides
        worksheet.conditional_format(first_row + 1,              first_col,
                                     first_row + rows_count - 2, first_col,
                                     {'type': 'formula', 'criteria': 'True', 'format': workbook.add_format({'left': thickness, 'right': thickness})})

        # bottom cap
        worksheet.conditional_format(first_row + rows_count - 1, first_col,
                                     first_row + rows_count - 1, first_col,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'bottom': thickness, 'left': thickness, 'right': thickness})})

    else:
        # top left corner
        worksheet.conditional_format(first_row, first_col,
                                     first_row, first_col,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'top': thickness, 'left': thickness})})

        # top right corner
        worksheet.conditional_format(first_row, first_col + cols_count - 1,
                                     first_row, first_col + cols_count - 1,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'top': thickness, 'right': thickness})})

        # bottom left corner
        worksheet.conditional_format(first_row + rows_count - 1, first_col,
                                     first_row + rows_count - 1, first_col,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'bottom': thickness, 'left': thickness})})

        # bottom right corner
        worksheet.conditional_format(first_row + rows_count - 1, first_col + cols_count - 1,
                                     first_row + rows_count - 1, first_col + cols_count - 1,
                                     {'type': 'formula', 'criteria': 'True',
                                      'format': workbook.add_format({'bottom': thickness, 'right': thickness})})

        # top
        worksheet.conditional_format(first_row, first_col + 1,
                                     first_row, first_col + cols_count - 2,
                                     {'type': 'formula', 'criteria': 'True', 'format': workbook.add_format({'top': thickness})})

        # left
        worksheet.conditional_format(first_row + 1,              first_col,
                                     first_row + rows_count - 2, first_col,
                                     {'type': 'formula', 'criteria': 'True', 'format': workbook.add_format({'left': thickness})})

        # bottom
        worksheet.conditional_format(first_row + rows_count - 1, first_col + 1,
                                     first_row + rows_count - 1, first_col + cols_count - 2,
                                     {'type': 'formula', 'criteria': 'True', 'format': workbook.add_format({'bottom': thickness})})

        # right
        worksheet.conditional_format(first_row + 1,              first_col + cols_count - 1,
                                     first_row + rows_count - 2, first_col + cols_count - 1,
                                     {'type': 'formula', 'criteria': 'True', 'format': workbook.add_format({'right': thickness})})


# Header
header = '&C&"Times New Roman"&B&U&20GST TAX INVOICE&18\n&U \n&UF.K. PATANWALA && Co.&U&18\n &14Hardware, Plumbing goods, Sanitary goods, Paints, Electrical materials && General merchant&18\n &14Address:- 67,Trinity Street, S.S Gaikwad Marg Dhobi Talao, Mumbai-400002.'


class App(Tk):
    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        container = Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.tk_setPalette(background=charcoal, foreground=navy,
                           activeBackground=rust, activeForeground=charcoal)
        self.geometry("1280x800+0+0")
        self.title("Billing Software")
        self.frames = {}
        for F in (Home, AddClient, CreateBill, AddBillDetails, EditBillDetails, UpdateBillStatus, GenerateBill, AddPurchaseBill, UpdatePurchaseStatus):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame(Home)

    def show_frame(self, page_name):
        '''Show a frame for the given page name'''
        frame = self.frames[page_name]
        frame.tkraise()


class Home(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.controller = controller

        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        saleF.place(x=1, rely=0.11, relwidth=1, relheight=0.45)

        # Create Sales Bill
        Button(saleF, text="Create New Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(CreateBill)).place(relx=0.15, rely=0.1, relwidth=0.3, relheight=0.2)

        # Add Sales Bill Button
        Button(saleF, text="Add Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddBillDetails)).place(relx=0.55, rely=0.1, relwidth=0.3, relheight=0.2)

        # Add Client Button
        Button(saleF, text="Add New Client", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddClient)).place(relx=0.15, rely=0.4, relwidth=0.3, relheight=0.2)

        # Edit Sales Bill Button
        Button(saleF, text="Edit Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(EditBillDetails)).place(relx=0.55, rely=0.4, relwidth=0.3, relheight=0.2)

        # Check & Edit Bill Status Button
        Button(saleF, text="Search & Edit Bill Status", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(UpdateBillStatus)).place(relx=0.15, rely=0.7, relwidth=0.3, relheight=0.2)

        # Generate Bill
        Button(saleF, text="Create & Print Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(GenerateBill)).place(relx=0.55, rely=0.7, relwidth=0.3, relheight=0.2)

        # --------- Purchase Options ---------
        purchaseF = LabelFrame(self, text="Purchase", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        purchaseF.place(x=1, rely=0.58, relwidth=1, relheight=0.22)

        # Create Purchase Bill
        Button(purchaseF, text="Create New Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddPurchaseBill)).place(relx=0.15, rely=0.2, relwidth=0.3, relheight=0.6)

        # Search Bill Button
        Button(purchaseF, text="Search & Edit Bill Status", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(UpdatePurchaseStatus)).place(relx=0.55, rely=0.2, relwidth=0.3, relheight=0.6)

        # --------- Backup Options ---------
        BackupF = LabelFrame(self, text="Backup", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"), padx=50, pady=15)
        BackupF.place(x=1, rely=0.82, relwidth=1, relheight=0.18)

        # Push to Git Button
        Button(BackupF, text="Upload to Cloud", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=self.gitPush).place(relx=0.3, rely=0.05, relwidth=0.4, relheight=0.9)

    def gitPush(self):
        cwd = os.getcwd()
        os.chdir(cwd + "/Bills")
        dt = datetime.now()
        message = "Backup [" + dt.strftime("%d-%m-%Y %I:%M %p")+"]"
        try:
            cmd.run("git pull origin master", check=True, shell=True)
            cmd.run("git add .", check=True, shell=True)
            cmd.run(f'git commit -m "{message}"', check=True, shell=True)
            cmd.run("git push -u origin master -f",
                    check=True, shell=True)
            messagebox.showinfo(
                title="Successful!!", message="Uploaded to Cloud: " + message)
        except cmd.SubprocessError as e:
            messagebox.showerror(
                title="Failed to Upload to Cloud!!!", message=e)

        finally:
            os.chdir(cwd)

# ---------- SALES FRAMES ----------


class AddClient(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.saleF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.saleF, text="Client's Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.04, relwidth=0.9, relheight=0.1)

        # Client's Name
        Label(self.saleF, text="Client's Name:", font=(
            "times new roman", 22, "bold"), pady=10).place(relx=0.1, rely=0.17, relwidth=0.3, relheight=0.1)
        self.cname_txt = Entry(self.saleF, width=40, bd=3, relief=GROOVE, font=(
            "arial", 22, "bold"))
        self.cname_txt.place(relx=0.45, rely=0.18,
                             relwidth=0.45, relheight=0.07)

        # Client's GST No
        Label(self.saleF, text="Client's GST No:", font=(
            "times new roman", 22, "bold"), pady=10).place(relx=0.1, rely=0.29, relwidth=0.3, relheight=0.1)
        self.cgst_txt = Entry(self.saleF, width=40, bd=3, relief=GROOVE, font=(
            "arial", 22, "bold"))
        self.cgst_txt.place(relx=0.45, rely=0.30,
                            relwidth=0.45, relheight=0.07)

        # Client's Address
        Label(self.saleF, text="Client's Address:", font=(
            "times new roman", 22, "bold"), pady=10).place(relx=0.1, rely=0.41, relwidth=0.3, relheight=0.1)
        self.caddress_txt = Entry(self.saleF, width=40, bd=3, relief=GROOVE, font=(
            "arial", 22, "bold"))
        self.caddress_txt.place(relx=0.45, rely=0.42,
                                relwidth=0.45, relheight=0.07)

        # Add Client Details Button
        add_client_btn = Button(self.saleF, text="Add New Client", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=50, font=(
            "arial", 18, "bold"), command=self.addClient).place(relx=0.2, rely=0.56, relwidth=0.6, relheight=0.1)

        # CLear Details Button
        clear_btn = Button(self.saleF, text="Clear", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=self.clearText).place(relx=0.2, rely=0.70, relwidth=0.6, relheight=0.1)

        # Home Button
        add_client_btn = Button(self.saleF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.2, rely=0.84, relwidth=0.6, relheight=0.1)

    def addClient(self):
        name = self.cname_txt.get()
        if not name:
            messagebox.showerror(
                title="Error", message="Name Field cannot be empty!!")
            return
        gst = self.cgst_txt.get()
        if not gst:
            messagebox.showerror(
                title="Error", message="GST Field cannot be empty!!")
            return
        address = self.caddress_txt.get()
        if not address:
            messagebox.showerror(
                title="Error", message="Address Field cannot be empty!!")
            return
        self.insertClient(name, address, gst)

    def insertClient(self, name, address, gst):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            sqlite_insert_with_param = """INSERT INTO client
                    (c_name, c_address, c_gst)
                    VALUES (?, ?, ?);"""

            data_tuple = (name, address, gst)
            cursor.execute(sqlite_insert_with_param, data_tuple)
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Client added successfully!!")
            getCientList()
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into client table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def clearText(self):
        self.cname_txt.delete(0, END)
        self.cgst_txt.delete(0, END)
        self.caddress_txt.delete(0, END)
        self.cname_txt.focus()


class CreateBill(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.saleF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.saleF, text="Bill Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.04, relwidth=0.9, relheight=0.1)

        # Billing Client's Name
        Label(self.saleF, text="Client's Name:", font=(
            "times new roman", 22, "bold")).place(relx=0.1, rely=0.17, relwidth=0.3, relheight=0.1)
        self.cname = StringVar()
        self.cname_txt = ttk.Combobox(self.saleF, width=40, font=(
            "arial", 22, "bold"), textvariable=self.cname, postcommand=self.updateClientList)
        self.cname_txt.place(relx=0.45, rely=0.18,
                             relwidth=0.45, relheight=0.07)

        # Billing Year
        Label(self.saleF, text="Financial Year:", font=(
            "times new roman", 22, "bold")).place(relx=0.1, rely=0.29, relwidth=0.3, relheight=0.1)
        self.byear = StringVar()
        self.byear_txt = OptionMenu(self.saleF, self.byear, *years)
        self.byear_txt.config(width=15, font=(
            "arial", 22, "bold"), bd=3, relief=GROOVE)
        self.byear_txt.place(relx=0.45, rely=0.30,
                             relwidth=0.45, relheight=0.07)

        # Billing Date
        Label(self.saleF, text="Date:", font=(
            "times new roman", 22, "bold")).place(relx=0.1, rely=0.41, relwidth=0.3, relheight=0.1)
        self.bdate = Entry(self.saleF, width=15, font=(
            "arial", 22, "bold"), bd=3, relief=GROOVE)
        self.bdate.place(relx=0.45, rely=0.42, relwidth=0.45, relheight=0.07)

        # P.O. Number
        Label(self.saleF, text="P.O. Number.:", font=(
            "times new roman", 22, "bold")).place(relx=0.1, rely=0.53, relwidth=0.3, relheight=0.1)
        self.bpo = Entry(self.saleF, width=15, font=(
            "arial", 22, "bold"), bd=3, relief=GROOVE)
        self.bpo.place(relx=0.45, rely=0.54, relwidth=0.45, relheight=0.07)

        # Add Bill Button
        Button(self.saleF, text="Add Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=50, font=(
            "arial", 18, "bold"), command=self.createBill).place(relx=0.2, rely=0.70, relwidth=0.6, relheight=0.1)

        # Add Bill Details Button
        clear_btn = Button(self.saleF, text="Add Bill Details", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddBillDetails)).place(relx=0.2, rely=0.84, relwidth=0.25, relheight=0.1)

        # Home Button
        add_client_btn = Button(self.saleF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.55, rely=0.84, relwidth=0.25, relheight=0.1)

    def createBill(self):
        name = self.cname_txt.get()
        if not name:
            messagebox.showerror(
                title="Error", message="Name Field cannot be empty!!")
            return
        if name not in clients:
            messagebox.showerror(
                title="Error", message="Select Client's Name from Dropdown List!!")
            return
        year = self.byear.get()
        if not year:
            messagebox.showerror(
                title="Error", message="Select a Year!!")
            return
        date = self.bdate.get()
        if not date:
            messagebox.showerror(
                title="Error", message="Enter a Date!!")
            return
        po = self.bpo.get()
        self.insertBill(name, year, date, po)

    def insertBill(self, name, year, date, po):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                """Select MAX(b_no) FROM bill WHERE b_year=?;""", (year,))
            val = cursor.fetchall()
            if val[0][0] == None:
                bill_no = 1
            else:
                bill_no = val[0][0]+1
            cursor.execute(
                """Select c_id FROM client WHERE c_name=?;""", (name,))
            val = cursor.fetchall()
            c_id = val[0][0]
            cursor.execute("""INSERT INTO  bill (b_year, b_no, c_id, b_date, b_po) VALUES (?,?,?,?,?);""",
                           (year, bill_no, c_id, date, po,))
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Bill added successfully with Bill No:"+str(bill_no)+"!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into bill table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def updateClientList(self):
        getCientList()
        self.cname_txt['values'] = clients


class AddBillDetails(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.saleF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.saleF, text="Bill Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.01, relwidth=0.9, relheight=0.07)

        # ---------- Add Details Frame ------------
        self.addDetailsF = LabelFrame(self.saleF, text="Add Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.addDetailsF.place(relx=0.01, rely=0.09,
                               relwidth=0.48, relheight=0.45)

        # Product Name
        Label(self.addDetailsF, text="Product:", font=(
            "times new roman", 18, "bold")).place(relx=0, rely=0.01, relwidth=0.2, relheight=0.1)
        self.pname = AutocompleteAddEntry(
            products, self.addDetailsF, listboxLength=15, matchesFunction=matches, bd=3, relief=GROOVE, font=("arial", 18, "bold"))
        self.pname.place(relx=0.2, rely=0.01, relwidth=0.75, relheight=0.13)

        # Quantity
        Label(self.addDetailsF, text="Quantity:", font=(
            "times new roman", 14, "bold")).place(relx=0, rely=0.18, relwidth=0.2, relheight=0.1)
        self.pquan = Entry(self.addDetailsF, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.pquan.place(relx=0.2, rely=0.18, relwidth=0.2, relheight=0.1)

        # Rate
        self.rate = DoubleVar()
        Label(self.addDetailsF, text="Rate:", font=(
            "times new roman", 14, "bold")).place(relx=0.55, rely=0.18, relwidth=0.2, relheight=0.1)
        self.rateE = Entry(self.addDetailsF, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"), textvariable=self.rate)
        self.rateE.place(relx=0.75, rely=0.18, relwidth=0.2, relheight=0.1)

        # GST
        self.gst = DoubleVar()
        # CGST
        Label(self.addDetailsF, text="GST:", font=(
            "times new roman", 14, "bold")).place(relx=0, rely=0.32, relwidth=0.2, relheight=0.1)
        self.gstE = Entry(self.addDetailsF, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"), textvariable=self.gst)
        self.gstE.place(relx=0.2, rely=0.32, relwidth=0.2, relheight=0.1)

        # IGST
        Label(self.addDetailsF, text="IGST:", font=(
            "times new roman", 14, "bold")).place(relx=0.55, rely=0.32, relwidth=0.2, relheight=0.1)
        self.igst = DoubleVar()
        self.igstE = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"), textvariable=self.igst)
        self.igstE.place(relx=0.75, rely=0.32, relwidth=0.2, relheight=0.1)

        # Challan No.
        Label(self.addDetailsF, text="Challan No:", font=(
            "times new roman", 14, "bold")).place(relx=0, rely=0.46, relwidth=0.2, relheight=0.1)
        self.challan = Entry(self.addDetailsF, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.challan.place(relx=0.2, rely=0.46, relwidth=0.2, relheight=0.1)

        # HSN Code
        Label(self.addDetailsF, text="HSN:", font=(
            "times new roman", 14, "bold")).place(relx=0.55, rely=0.46, relwidth=0.2, relheight=0.1)
        self.phsn = Entry(self.addDetailsF, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.phsn.place(relx=0.75, rely=0.46, relwidth=0.2, relheight=0.1)

        # AMC No.
        Label(self.addDetailsF, text="AMC No:", font=(
            "times new roman", 14, "bold")).place(relx=0.3, rely=0.60, relwidth=0.2, relheight=0.1)
        self.amc = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.amc.place(relx=0.5, rely=0.60, relwidth=0.2, relheight=0.1)

        # Insert Bill Details Button
        Button(self.addDetailsF, text="Add Bill Detail", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 16, "bold"), command=self.addBillDetails).place(relx=0.1, rely=0.76, relwidth=0.5, relheight=0.14)

        # Clear Details Button
        Button(self.addDetailsF, text="Clear", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 16, "bold"), command=self.clearTextAdd).place(relx=0.7, rely=0.76, relwidth=0.2, relheight=0.14)

        # --------- Select Bill Frame ------------
        self.SelectBillF = LabelFrame(self.saleF, text="Select Bill", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=20, pady=10)
        self.SelectBillF.place(relx=0.51, rely=0.09,
                               relwidth=0.48, relheight=0.45)

        # Billing Year
        Label(self.SelectBillF, text="Billing Year:", font=(
            "times new roman", 16, "bold")).place(relx=0, rely=0.01, relwidth=0.4, relheight=0.15)
        self.byear = StringVar()
        self.byearC = ttk.Combobox(self.SelectBillF, font=(
            "arial", 16, "bold"), textvariable=self.byear, values=years)
        self.byearC.place(relx=0.4, rely=0.01, relwidth=0.4, relheight=0.15)

        # Bill No.
        Label(self.SelectBillF, text="Billing No:", font=(
            "times new roman", 16, "bold")).place(relx=0, rely=0.3, relwidth=0.4, relheight=0.15)
        self.bno = StringVar()
        self.bnoC = Entry(self.SelectBillF, bd=3, relief=GROOVE, font=(
            "arial", 16, "bold"), textvariable=self.bno)
        self.bnoC.place(relx=0.4, rely=0.3, relwidth=0.4, relheight=0.15)

        # Current Bill Details Button
        Button(self.SelectBillF, text="View Bill Details", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 16, "bold"), command=self.viewBillDetails).place(relx=0.1, rely=0.61, relwidth=0.35, relheight=0.15)

        # Delete Last Entry Bill Details Button
        Button(self.SelectBillF, text="Delete Last Entry", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 16, "bold"), command=self.removeLastDetail).place(relx=0.55, rely=0.61, relwidth=0.35, relheight=0.15)

        # Home Button
        Button(self.SelectBillF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=35, pady=30, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.3, rely=0.81, relwidth=0.4, relheight=0.15)

        # --------- View Bill Detials Frame ------------
        self.viewBillF = LabelFrame(self.saleF, text="View Bill Detials", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.viewBillF.place(relx=0.01, rely=0.54,
                             relwidth=0.98, relheight=0.45)
        self.displayBillF = Frame(self.viewBillF)
        self.displayBillF.place(relx=0.01, rely=0.01,
                                relwidth=0.98, relheight=0.98)
        self.displayText = scrolledtext.ScrolledText(
            self.displayBillF, font=("Courier",
                                     12, "bold"), padx=10, pady=10)
        self.displayText.insert(INSERT,
                                "\n\n\n\n\n\n\n\n\t\t\t\t\tSelect Billing Year and Billing No. to see the entries!! ")
        self.displayText.configure(state='disabled')
        self.displayText.pack(side="left", fill="both", expand=True)
        Label(self.viewBillF, text="Sr. No.\t\t            Product\t\t              Rate      Quantity  GST     IGST       AMOUNT\t        HSN       Ch. No.    AMC No.", font=(
            "times new roman", 15, "bold")).place(x=30, y=0)

    def addBillDetails(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        name = self.pname.get()
        if not name:
            messagebox.showerror(
                title="Error", message="Product Field cannot be empty!!")
            return
        quantity = self.pquan.get()
        if not quantity:
            messagebox.showerror(
                title="Error", message="Quantity Field cannot be empty!!")
            return
        if not str.isdigit(quantity):
            messagebox.showerror(
                title="Error", message="Enter a Valid Quantity!!")
            return
        rate = str(self.rate.get())
        if not rate:
            messagebox.showerror(
                title="Error", message="Rate Field cannot be empty!!")
            return
        amt = int(Decimal(float(rate) * int(quantity)).quantize(0, ROUND_HALF_UP))
        gst = self.gst.get() * 0.5
        igst = self.igst.get()
        if not gst and not igst:
            messagebox.showerror(
                title="Error", message="Enter CGST/SGST or IGST!!")
            return
        hsn = self.phsn.get()
        amc = self.amc.get()
        challan = self.challan.get()

        data_list = [b_year, b_no, name, rate, quantity,
                     gst, igst, amt, hsn, challan, amc]
        self.insertBillDetails(data_list)

    def insertBillDetails(self, data_list):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select b_no FROM bill WHERE b_year=? AND b_no=?;", (data_list[0], data_list[1],))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            cursor.execute(
                """Select MAX(bd_id) FROM bill_detail WHERE b_year=? AND b_no=?;""", (data_list[0], data_list[1],))
            val = cursor.fetchall()
            if val[0][0] == None:
                id = 1
            else:
                id = val[0][0]+1
            data_list.insert(2, id)
            sqlite_insert_with_param = """INSERT INTO bill_detail
                            (b_year, b_no,bd_id, bd_product, bd_rate, bd_quantity,
                             bd_cgst, bd_igst, bd_amount, bd_hsn, bd_ch_no, bd_amc_no)
                            VALUES (?, ?, ?,?, ?, ?,?, ?,?,?,?, ?);"""
            cursor.execute(sqlite_insert_with_param, tuple(data_list))
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Details added successfully!!")
            cursor.close()
            self.viewBillDetails()
            if data_list[3] not in products:
                addProduct(data_list[3])
            self.clearTextAdd()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into bill_detail table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def viewBillDetails(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        self.SelectBillDetails(b_year, b_no)

    def SelectBillDetails(self, b_year, b_no):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            sqlite_insert_with_param = """Select bd_id, bd_product, bd_rate, bd_quantity, bd_cgst, bd_igst, bd_amount, bd_hsn, bd_ch_no, bd_amc_no FROM bill_detail WHERE b_year=? AND b_no=?;"""

            data_tuple = (b_year, b_no)
            cursor.execute(sqlite_insert_with_param, data_tuple)
            rows = cursor.fetchall()
            if not rows:
                cursor.execute(
                    "Select b_no FROM bill WHERE b_year=? AND b_no=?", data_tuple)
                rows = cursor.fetchall()
                if not rows:
                    messagebox.showerror(
                        title="Error", message="Bill No. does not EXIST!!")
                    return
                else:
                    messagebox.showerror(
                        title="Error", message="Bill No. has no Entries!!")
                    return
            s = ""
            for row in rows:
                s += "\n" + str(row[0]).center(7)
                if len(row[1]) < 36:
                    s += str(row[1]).center(38)
                else:
                    s += str(row[1])[:36]+".."
                s += str(row[2]).center(10)  # Rate
                s += str(row[3]).center(7)  # Quantity
                s += str(row[4]*2).center(7)  # GST
                s += str(row[5]).center(7)  # IGST
                s += str(row[6]).center(14)  # Amount
                s += str(row[7]).center(10)  # HSN
                s += str(row[8]).center(7)  # Challan
                s += str(row[9]).center(10)  # AMC
            self.displayText.configure(state='normal')
            self.displayText.delete('1.0', END)
            self.displayText.insert(INSERT, s)
            self.displayText.configure(state='disabled')
            sqliteConnection.commit()
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get data from bill_detail table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def removeLastDetail(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        self.deleteDetails(b_year, b_no)

    def deleteDetails(self, b_year, b_no):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select b_no FROM bill WHERE b_year=? AND b_no=?;", (b_year, b_no,))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            cursor.execute(
                "Select MAX(bd_id) FROM bill_detail WHERE b_year=? AND b_no=?;", (b_year, b_no,))
            rows = cursor.fetchall()
            if not rows[0][0]:
                messagebox.showerror(
                    title="Error", message="Bill has No Entries!!")
                return
            cursor.execute(
                "DELETE from bill_detail WHERE bd_id = ? AND b_year = ? AND b_no = ?;", (rows[0][0], b_year, b_no,))
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Details Deleted Successfully!!")
            getCientList()
            cursor.close()
            self.viewBillDetails()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into bill_detail table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def clearTextAdd(self):
        self.pname.delete(0, END)
        self.rateE.delete(0, END)
        self.pquan.delete(0, END)
        self.gstE.delete(0, END)
        self.gstE.insert(0,0)
        self.igstE.delete(0, END)
        self.igstE.insert(0,0)
        self.phsn.delete(0, END)
        self.challan.delete(0, END)
        self.amc.delete(0, END)
        self.pname.focus()

    def updateClientList(self):
        getCientList()
        self.cname_txt['values'] = clients


class EditBillDetails(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.saleF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.saleF, text="Bill Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.01, relwidth=0.9, relheight=0.07)

        # --------- Edit Details Frame ------------
        self.editDetailsF = LabelFrame(self.saleF, text="Edit Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.editDetailsF.place(relx=0.01, rely=0.09,
                                relwidth=0.38, relheight=0.45)

        # Bill Serial No.
        Label(self.editDetailsF, text="Sr. No.:", font=(
            "times new roman", 14, "bold")).place(relx=0.01, rely=0.01, relwidth=0.16, relheight=0.1)
        self.esrno = StringVar()
        self.esrnoC = Entry(self.editDetailsF, width=4, font=(
            "arial", 14, "bold"), textvariable=self.esrno, bd=3, relief=GROOVE)
        self.esrnoC.place(relx=0.16, rely=0.01, relwidth=0.16, relheight=0.1)

        # HSN Code
        Label(self.editDetailsF, text="HSN:", font=(
            "times new roman", 14, "bold")).place(relx=0.32, rely=0.01, relwidth=0.16, relheight=0.1)
        self.ephsn = Entry(self.editDetailsF, width=6, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.ephsn.place(relx=0.48, rely=0.01, relwidth=0.16, relheight=0.1)

        # AMC No.
        Label(self.editDetailsF, text="AMC:", font=(
            "times new roman", 14, "bold")).place(relx=0.64, rely=0.01, relwidth=0.16, relheight=0.1)
        self.eamc = Entry(self.editDetailsF, width=7, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.eamc.place(relx=0.8, rely=0.01, relwidth=0.16, relheight=0.1)

        # Product Name
        Label(self.editDetailsF, text="Product:", font=(
            "times new roman", 14, "bold")).place(relx=0, rely=0.16, relwidth=0.2, relheight=0.1)
        self.epname = AutocompleteEntry(
            products, self.editDetailsF, listboxLength=10, matchesFunction=matches, width=15, bd=3, relief=GROOVE, font=("arial", 14, "bold"))
        self.epname.place(relx=0.2, rely=0.16, relwidth=0.75, relheight=0.1)

        # Quantity
        Label(self.editDetailsF, text="Quantity:", font=(
            "times new roman", 14, "bold")).place(relx=0, rely=0.31, relwidth=0.2, relheight=0.1)
        self.epquan = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.epquan.place(relx=0.2, rely=0.31, relwidth=0.2, relheight=0.1)

        # Rate
        Label(self.editDetailsF, text="Rate:", font=(
            "times new roman", 14, "bold")).place(relx=0.55, rely=0.31, relwidth=0.2, relheight=0.1)
        self.erate = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.erate.place(relx=0.75, rely=0.31, relwidth=0.2, relheight=0.1)

        # Taxable Amount
        Label(self.editDetailsF, text="Amount:", font=(
            "times new roman", 14, "bold")).place(relx=0, rely=0.46, relwidth=0.2, relheight=0.1)
        self.eamt = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.eamt.place(relx=0.2, rely=0.46, relwidth=0.2, relheight=0.1)

        # GST
        Label(self.editDetailsF, text="GST:", font=(
            "times new roman", 14, "bold")).place(relx=0.55, rely=0.46, relwidth=0.2, relheight=0.1)
        self.ecgst = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.ecgst.place(relx=0.75, rely=0.46, relwidth=0.2, relheight=0.1)

        # IGST
        Label(self.editDetailsF, text="IGST:", font=(
            "times new roman", 14, "bold")).place(relx=0, rely=0.61, relwidth=0.2, relheight=0.1)
        self.eigst = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.eigst.place(relx=0.2, rely=0.61, relwidth=0.2, relheight=0.1)

        # Challan No.
        Label(self.editDetailsF, text="Challan No:", font=(
            "times new roman", 14, "bold")).place(relx=0.5, rely=0.61, relwidth=0.2, relheight=0.1)
        self.echallan = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.echallan.place(relx=0.75, rely=0.61, relwidth=0.2, relheight=0.1)

        # Edit Bill Details Button
        Button(self.editDetailsF, text="Edit Bill Detail", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=15, font=(
            "arial", 16, "bold"), command=self.editBillDetails).place(relx=0.1, rely=0.8, relwidth=0.35, relheight=0.15)

        # Clear Edit Details Button
        Button(self.editDetailsF, text="Clear", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=15, font=(
            "arial", 16, "bold"), command=self.clearTextEdit).place(relx=0.55, rely=0.8, relwidth=0.35, relheight=0.15)

        # --------- Select Bill Frame ------------
        self.SelectBillF = LabelFrame(self.saleF, text="Select Bill", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.SelectBillF.place(relx=0.4, rely=0.09,
                               relwidth=0.25, relheight=0.45)

        # Billing Year
        Label(self.SelectBillF, text="Billing Year:", font=(
            "times new roman", 16, "bold")).place(relx=0, rely=0, relwidth=0.45, relheight=0.15)
        self.byear = StringVar()
        self.byearC = ttk.Combobox(self.SelectBillF, font=(
            "arial", 16, "bold"), textvariable=self.byear, values=years)
        self.byearC.place(relx=0.45, rely=0.01, relwidth=0.5, relheight=0.12)

        # Bill No.
        Label(self.SelectBillF, text="Billing No:", font=(
            "times new roman", 16, "bold")).place(relx=0, rely=0.17, relwidth=0.45, relheight=0.15)
        self.bno = StringVar()
        self.bnoC = Entry(self.SelectBillF, bd=3, relief=GROOVE, font=(
            "arial", 16, "bold"), textvariable=self.bno)
        self.bnoC.place(relx=0.45, rely=0.19, relwidth=0.5, relheight=0.12)

        # Current Bill Details Button
        Button(self.SelectBillF, text="View Bill Details", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 16, "bold"), command=self.viewBillDetails).place(relx=0.15, rely=0.4, relwidth=0.7, relheight=0.15)

        # Delete Last Entry Bill Details Button
        Button(self.SelectBillF, text="Delete Last Entry", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 16, "bold"), command=self.removeLastDetail).place(relx=0.15, rely=0.6, relwidth=0.7, relheight=0.15)

        # Home Button
        Button(self.SelectBillF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=35, pady=30, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.15, rely=0.8, relwidth=0.7, relheight=0.15)

        # --------- Edit Client Frame ------------
        self.editClientF = LabelFrame(self.saleF, text="Edit Client Name", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.editClientF.place(relx=0.66, rely=0.09,
                               relwidth=0.33, relheight=0.15)

        # Billing Client's Name
        self.ecname = StringVar()
        self.cname_txt = ttk.Combobox(self.editClientF, width=25, font=(
            "arial", 16, "bold"), textvariable=self.ecname, postcommand=self.updateClientList)
        self.cname_txt.place(relx=0.05, rely=0.3, relwidth=0.5, relheight=0.4)

        # Edit Client Button
        Button(self.editClientF, text="Edit Client", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=15, pady=10, font=(
            "arial", 16, "bold"), command=self.editClient).place(relx=0.6, rely=0.22, relwidth=0.35, relheight=0.56)

        # --------- Bill Information Frame ------------
        self.viewInfoF = LabelFrame(self.saleF, text="Bill Information", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.viewInfoF.place(relx=0.66, rely=0.25,
                             relwidth=0.33, relheight=0.29)
        self.displayInfoF = Frame(self.viewInfoF)
        self.displayInfoF.place(
            relx=0.02, rely=0.05, relwidth=0.96, relheight=0.9)
        self.displayInfo = scrolledtext.ScrolledText(
            self.displayInfoF, font=("arial",
                                     16, "bold"), padx=10, pady=10)
        self.displayInfo.insert(INSERT,
                                "\nSelect Billing Year and Billing No. \nto see the Information!! ")
        self.displayInfo.configure(state='disabled')
        self.displayInfo.pack(side="left", fill="both", expand=True)

        # --------- View Bill Detials Frame ------------
        self.viewBillF = LabelFrame(self.saleF, text="View Bill Detials", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.viewBillF.place(relx=0.01, rely=0.54,
                             relwidth=0.98, relheight=0.45)
        self.displayBillF = Frame(self.viewBillF)
        self.displayBillF.place(relx=0.01, rely=0.01,
                                relwidth=0.98, relheight=0.98)
        self.displayText = scrolledtext.ScrolledText(
            self.displayBillF, font=("Courier",
                                     12, "bold"), padx=10, pady=10)
        self.displayText.insert(INSERT,
                                "\n\n\n\n\n\n\n\n\t\t\t\t\tSelect Billing Year and Billing No. to see the entries!! ")
        self.displayText.configure(state='disabled')
        self.displayText.pack(side="left", fill="both", expand=True)
        Label(self.viewBillF, text="Sr. No.\t\t            Product\t\t              Rate      Quantity  GST     IGST       AMOUNT\t        HSN       Ch. No.    AMC No.", font=(
            "times new roman", 15, "bold")).place(x=30, y=0)

    def editBillDetails(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        constraints = []
        srno = self.esrno.get()
        if not srno:
            messagebox.showerror(
                title="Error", message="Serial No. Field cannot be empty!!")
            return
        name = self.epname.get()
        if name:
            if name not in products:
                messagebox.showerror(
                    title="Error", message="Select Product's Name from Dropdown List!!")
                return
            constraints.append("bd_product")
            constraints.append(name)
        rate = str(self.erate.get())
        amt = self.eamt.get()
        quantity = self.epquan.get()
        if rate or amt or quantity:
            if not str.isdigit(quantity):
                messagebox.showerror(
                    title="Error", message="Enter a Valid Quantity!!")
                return
            if not rate:
                messagebox.showerror(
                    title="Error", message="Enter a Valid Rate!!")
                return
            amt = int(Decimal(float(rate) * int(quantity)).quantize(0, ROUND_HALF_UP))
            constraints.append("bd_rate")
            constraints.append(rate)
            constraints.append("bd_amount")
            constraints.append(str(amt))
            constraints.append("bd_quantity")
            constraints.append(quantity)
        cgst = self.ecgst.get()
        if cgst:
            constraints.append("bd_cgst")
            constraints.append(cgst)
        igst = self.eigst.get()
        if igst:
            constraints.append("bd_igst")
            constraints.append(igst)
        hsn = self.ephsn.get()
        if hsn:
            constraints.append("bd_hsn")
            constraints.append(hsn)
        amc = self.eamc.get()
        if amc:
            constraints.append("bd_amc_no")
            constraints.append(amc)
        challan = self.echallan.get()
        if challan:
            constraints.append("bd_ch_no")
            constraints.append(challan)
        if not len(constraints):
            messagebox.showerror(
                title="Error", message="Enter Atleast 1 value to be updated!!")
            return
        self.updateBill(b_year, b_no, srno, constraints)

    def updateBill(self, b_year, b_no, srno, constraints):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select b_no FROM bill WHERE b_year=? AND b_no=?;", (b_year, b_no,))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            data_tuple = (srno, b_year, b_no,)
            cursor.execute(
                "Select bd_id FROM bill_detail WHERE bd_id=? AND b_year=? AND b_no=?;", data_tuple)
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Sr. No. does not EXIST!!")
                return
            for i in range(0, len(constraints), 2):
                sqlite_insert_with_param = "UPDATE bill_detail SET " + \
                    constraints[i] + " = \"" + constraints[i+1] + \
                    "\" WHERE bd_id = ? AND b_year = ? AND b_no = ?;"
                cursor.execute(
                    sqlite_insert_with_param, data_tuple)
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Details added successfully!!")
            getCientList()
            cursor.close()
            self.viewBillDetails()
            if "bd_product" == constraints[0] and constraints[1] not in products:
                addProduct(constraints[1])

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into Bill Details table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def viewBillDetails(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        self.SelectBillDetails(b_year, b_no)
        self.SelectBillInfo(b_year, b_no)

    def SelectBillDetails(self, b_year, b_no):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            sqlite_insert_with_param = """Select bd_id, bd_product, bd_rate, bd_quantity, bd_cgst, bd_igst, bd_amount, bd_hsn, bd_ch_no, bd_amc_no FROM bill_detail WHERE b_year=? AND b_no=?;"""

            data_tuple = (b_year, b_no)
            cursor.execute(sqlite_insert_with_param, data_tuple)
            rows = cursor.fetchall()
            if not rows:
                cursor.execute(
                    "Select b_no FROM bill WHERE b_year=? AND b_no=?", data_tuple)
                rows = cursor.fetchall()
                if not rows:
                    messagebox.showerror(
                        title="Error", message="Bill No. does not EXIST!!")
                    return
                else:
                    messagebox.showerror(
                        title="Error", message="Bill No. has no Entries!!")
                    return
            s = ""
            for row in rows:
                s += "\n" + str(row[0]).center(7)
                if len(row[1]) < 36:
                    s += str(row[1]).center(38)
                else:
                    s += str(row[1])[:36]+".."
                s += str(row[2]).center(10)  # Rate
                s += str(row[3]).center(7)  # Quantity
                s += str(row[4]*2).center(7)  # GST
                s += str(row[5]).center(7)  # IGST
                s += str(row[6]).center(14)  # Amount
                s += str(row[7]).center(10)  # HSN
                s += str(row[8]).center(7)  # Challan
                s += str(row[9]).center(10)  # AMC
            self.displayText.configure(state='normal')
            self.displayText.delete('1.0', END)
            self.displayText.insert(INSERT, s)
            self.displayText.configure(state='disabled')
            sqliteConnection.commit()
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get data from bill_detail table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def removeLastDetail(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        self.deleteDetails(b_year, b_no)

    def deleteDetails(self, b_year, b_no):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select b_no FROM bill WHERE b_year=? AND b_no=?;", (b_year, b_no,))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            cursor.execute(
                "Select MAX(bd_id) FROM bill_detail WHERE b_year=? AND b_no=?;", (b_year, b_no,))
            rows = cursor.fetchall()
            if not rows[0][0]:
                messagebox.showerror(
                    title="Error", message="Bill has No Entries!!")
                return
            cursor.execute(
                "DELETE from bill_detail WHERE bd_id = ? AND b_year = ? AND b_no = ?;", (rows[0][0], b_year, b_no,))
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Details Deleted Successfully!!")
            getCientList()
            cursor.close()
            self.viewBillDetails()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into bill_detail table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def editClient(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        cname = self.ecname.get()
        if cname not in clients:
            messagebox.showerror(
                title="Error", message="Select Client Name from Dropdown List!!")
            return
        self.updateClient(b_year, b_no, cname)

    def updateClient(self, b_year, b_no, cname):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute("Select c_id FROM client WHERE c_name=?;", (cname,))
            c_id = cursor.fetchone()
            sqlite_insert_with_param = "UPDATE bill SET c_id = ? WHERE b_year = ? AND b_no = ?;"
            cursor.execute(
                sqlite_insert_with_param, (c_id[0], b_year, b_no,))
            sqliteConnection.commit()
            getCientList()
            cursor.close()
            self.SelectBillInfo(b_year, b_no)

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into Bill table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def SelectBillInfo(self, b_year, b_no):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            sqlite_insert_with_param = """Select c_id FROM bill WHERE b_year=? AND b_no=?;"""
            data_tuple = (b_year, b_no)
            cursor.execute(sqlite_insert_with_param, data_tuple)
            rows = cursor.fetchall()
            if not rows:
                return
            c_id = rows[0][0]
            sqlite_insert_with_param = """Select c_name FROM client WHERE c_id=?;"""
            cursor.execute(sqlite_insert_with_param, (c_id,))
            rows = cursor.fetchall()
            s = "Client's Name: "+rows[0][0]+"\n\n"
            sqlite_insert_with_param = """Select b_date, b_status, b_po FROM bill WHERE b_year=? AND b_no=?;"""
            cursor.execute(sqlite_insert_with_param, (b_year, b_no,))
            rows = cursor.fetchall()
            s += "Date: "+rows[0][0]+"\n\n"
            s += "Payment Status: "+rows[0][1]+"\n\n"
            s += "P.O. Number: "+rows[0][2]
            self.displayInfo.configure(state='normal')
            self.displayInfo.delete('1.0', END)
            self.displayInfo.insert(INSERT, s)
            self.displayInfo.configure(state='disabled')
            sqliteConnection.commit()
            getCientList()
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to Get Data form Database", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def clearTextEdit(self):
        self.esrnoC.delete(0, END)
        self.epname.delete(0, END)
        self.ephsn.delete(0, END)
        self.epquan.delete(0, END)
        self.erate.delete(0, END)
        self.eamt.delete(0, END)
        self.ecgst.delete(0, END)
        self.eigst.delete(0, END)
        self.echallan.delete(0, END)
        self.eamc.delete(0, END)
        self.esrnoC.focus()

    def updateClientList(self):
        getCientList()
        self.cname_txt['values'] = clients


class UpdateBillStatus(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.saleF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.saleF, text="Check & Edit Bill Payment", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.03, relwidth=0.9, relheight=0.09)

        # --------- Check Bill Status Frame ------------
        self.SelectBillF = LabelFrame(self.saleF, text="  Search Bill Payment  ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"))
        self.SelectBillF.place(relx=0.05, rely=0.14,
                               relwidth=0.43, relheight=0.38)

        # Billing Year
        Label(self.SelectBillF, text="Billing Year:", font=(
            "times new roman", 18, "bold")).place(relx=0.1, rely=0.1, relwidth=0.3, relheight=0.12)
        self.byear = StringVar()
        self.byearC = ttk.Combobox(self.SelectBillF, font=(
            "arial", 18, "bold"), textvariable=self.byear, values=years)
        self.byearC.place(relx=0.4, rely=0.11, relwidth=0.5, relheight=0.11)

        # Client's Name
        Label(self.SelectBillF, text="Client's Name:", font=(
            "times new roman", 18, "bold")).place(relx=0.1, rely=0.35, relwidth=0.3, relheight=0.12)
        self.cname = StringVar()
        self.cnameC = ttk.Combobox(self.SelectBillF, font=(
            "arial", 18, "bold"), textvariable=self.cname, postcommand=self.updateClientList)
        self.cnameC.place(relx=0.4, rely=0.35, relwidth=0.5, relheight=0.14)

        # Search Bills Button
        Button(self.SelectBillF, text="Search Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=self.searchBill).place(relx=0.1, rely=0.65, relwidth=0.35, relheight=0.21)

        # Home Button
        Button(self.SelectBillF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.55, rely=0.65, relwidth=0.35, relheight=0.21)

        # --------- Edit Bill Status Frame ------------
        self.editStatusF = LabelFrame(self.saleF, text="  Update Bill Payment  ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"))
        self.editStatusF.place(relx=0.52, rely=0.14,
                               relwidth=0.43, relheight=0.38)

        # Bill No.
        Label(self.editStatusF, text="Billing No:", font=(
            "times new roman", 18, "bold")).place(relx=0.1, rely=0.1, relwidth=0.4, relheight=0.12)
        self.bno = StringVar()
        self.bnoC = Entry(self.editStatusF, width=30, font=(
            "arial", 18, "bold"), textvariable=self.bno, bd=3, relief=GROOVE)
        self.bnoC.place(relx=0.55, rely=0.11, relwidth=0.4, relheight=0.14)

        # Payment Status
        Label(self.editStatusF, text="Payment Status:", font=(
            "times new roman", 18, "bold")).place(relx=0.1, rely=0.35, relwidth=0.4, relheight=0.12)
        self.estatus = Entry(self.editStatusF, width=30, font=(
            "arial", 18, "bold"), bd=3, relief=GROOVE)
        self.estatus.place(relx=0.55, rely=0.35, relwidth=0.4, relheight=0.14)

        # Current Bill Details Button
        Button(self.editStatusF, text="Update Status", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=self.editStatus).place(relx=0.3, rely=0.65, relwidth=0.4, relheight=0.21)

        # --------- View Bill Detials Frame ------------
        self.viewBillF = LabelFrame(self.saleF, text="View Bill Detials", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=10, pady=10)
        self.viewBillF.place(relx=0.05, rely=0.54,
                             relwidth=0.9, relheight=0.45)
        self.displayBillF = Frame(self.viewBillF)
        self.displayBillF.place(relx=0.01, rely=0.05,
                                relwidth=0.98, relheight=0.98)
        self.displayText = scrolledtext.ScrolledText(
            self.displayBillF, font=("Courier",
                                     12, "bold"), padx=10, pady=10)
        self.displayText.insert(INSERT,
                                "\n\n\n\n\n\n\t\t\tSelect Billing Year or Client's Name to see the entries!! ")
        self.displayText.configure(state='disabled')
        self.displayText.pack(side="left", fill="both", expand=True)
        Label(self.viewBillF, text="Bill. No.\t       Year\t\t\t\t     Client's Name\t\t\t              Payment", font=(
            "times new roman", 15, "bold")).place(x=30, y=0)

    def searchBill(self):
        constraints = []
        c_name = self.cname.get()
        if c_name:
            if c_name not in clients:
                messagebox.showerror(
                    title="Error", message="Select Client's Name from Dropdown List!!")
                return
            constraints.append("c_name")
            constraints.append(c_name)
        b_year = self.byear.get()
        if b_year:
            if b_year not in years:
                messagebox.showerror(
                    title="Error", message="Select Blling Year from Dropdown List!!")
                return
            constraints.append(b_year)
        if not len(constraints):
            messagebox.showerror(
                title="Error", message="Select atleast 1 critera to search!!")
            return
        self.SelectBill(constraints)

    def SelectBill(self, constraints):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            s = ""
            if constraints[0] == "c_name":
                cursor.execute(
                    "Select c_id FROM client WHERE c_name =?;", (constraints[1],))
                rows = cursor.fetchone()
                c_id = rows[0]
                if len(constraints) == 2:
                    cursor.execute(
                        "Select b_no, b_year, b_status FROM bill WHERE c_id=?;", (c_id,))
                    rows = cursor.fetchall()
                else:
                    cursor.execute("Select b_no, b_year, b_status FROM bill WHERE c_id=? AND b_year=?;",
                                   (c_id, constraints[2],))
                    rows = cursor.fetchall()
                for row in rows:
                    s += "\n"+str(row[0]).center(7)
                    s += str(row[1]).center(17)
                    s += constraints[1].center(51)
                    s += str(row[2]).center(17)
            else:
                cursor.execute("Select b_no, b_year,c_id, b_status FROM bill WHERE b_year=?;",
                               (constraints[0],))
                rows = cursor.fetchall()
                for row in rows:
                    s += "\n"+str(row[0]).center(7)
                    s += str(row[1]).center(17)
                    cursor.execute(
                        "Select c_name FROM client WHERE c_id =?;", (row[2],))
                    c_name = cursor.fetchall()
                    s += c_name[0][0].center(51)
                    s += str(row[3]).center(17)
            self.displayText.configure(state='normal')
            self.displayText.delete('1.0', END)
            self.displayText.insert(INSERT, s)
            self.displayText.configure(state='disabled')
            sqliteConnection.commit()
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to Get Data form Database", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def editStatus(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        status = self.estatus.get()
        if not status:
            messagebox.showerror(
                title="Error", message="Payment Status Field cannot be empty!!")
            return
        self.updateStatus((status, b_year, b_no,))

    def updateStatus(self, constraints):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute("Select b_no FROM bill WHERE b_year=? AND b_no=?", (
                constraints[1], constraints[2],))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            cursor.execute(
                "UPDATE bill SET b_status=? WHERE b_year=? AND b_no=?;", constraints)
            messagebox.showinfo(
                title="Successfull", message="Payment Status updated Successfully!!")
            sqliteConnection.commit()
            cursor.close()
            self.SelectBill([constraints[1]])

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to Get Data form Database", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def updateClientList(self):
        getCientList()
        self.cnameC['values'] = clients


class GenerateBill(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.saleF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.saleF, text="Bill Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.03, relwidth=0.9, relheight=0.09)

        # Financial Year
        Label(self.saleF, text="Financial Year:", font=(
            "times new roman", 22, "bold")).place(relx=0.1, rely=0.17, relwidth=0.3, relheight=0.12)
        self.byear = StringVar()
        self.byearC = ttk.Combobox(self.saleF, width=40, font=(
            "arial", 22, "bold"), textvariable=self.byear, values=years)
        self.byearC.place(relx=0.4, rely=0.2, relwidth=0.5, relheight=0.06)

        # Bill No
        Label(self.saleF, text="Bill No.:", font=(
            "times new roman", 22, "bold"), pady=10).place(relx=0.1, rely=0.30, relwidth=0.3, relheight=0.12)
        self.billno = Entry(self.saleF, width=40, bd=3, relief=GROOVE, font=(
            "arial", 22, "bold"))
        self.billno.place(relx=0.4, rely=0.32, relwidth=0.5, relheight=0.07)

        # Bill Type
        Label(self.saleF, text="Bill Type:", font=(
            "times new roman", 22, "bold")).place(relx=0.1, rely=0.43, relwidth=0.3, relheight=0.12)
        self.btype = IntVar()
        Radiobutton(self.saleF, text="Normal", variable=self.btype, value=1, font=(
            "times new roman", 16, "bold"), bg=rust, bd=3, relief=GROOVE,
            indicatoron=0).place(relx=0.4, rely=0.45, relwidth=0.2, relheight=0.07)
        Radiobutton(self.saleF, text="Normal+AMC", variable=self.btype, value=2, font=(
            "times new roman", 16, "bold"), bg=rust, bd=3, relief=GROOVE,
            indicatoron=0).place(relx=0.7, rely=0.45, relwidth=0.2, relheight=0.07)
        Radiobutton(self.saleF, text="Normal + HSN", variable=self.btype, value=3, font=(
            "times new roman", 16, "bold"), bg=rust, bd=3, relief=GROOVE,
            indicatoron=0).place(relx=0.4, rely=0.55, relwidth=0.2, relheight=0.07)
        Radiobutton(self.saleF, text="Normal + IGST", variable=self.btype, value=4, font=(
            "times new roman", 16, "bold"), bg=rust, bd=3, relief=GROOVE,
            indicatoron=0).place(relx=0.7, rely=0.55, relwidth=0.2, relheight=0.07)
        Radiobutton(self.saleF, text="AMC + HSN", variable=self.btype, value=5, font=(
            "times new roman", 16, "bold"), bg=rust, bd=3, relief=GROOVE,
            indicatoron=0).place(relx=0.55, rely=0.65, relwidth=0.2, relheight=0.07)

        # Generate Excel Button
        add_client_btn = Button(self.saleF, text="Create Excel Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=50, font=(
            "arial", 18, "bold"), command=self.getBillData).place(relx=0.1, rely=0.78, relwidth=0.35, relheight=0.08)
        # Print Excel Button
        add_client_btn = Button(self.saleF, text="Print Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=50, font=(
            "arial", 18, "bold"), command=self.printBill).place(relx=0.55, rely=0.78, relwidth=0.35, relheight=0.08)
        # Home Button
        add_client_btn = Button(self.saleF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=50, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.3, rely=0.88, relwidth=0.4, relheight=0.08)

    def getBillData(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Financial Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.billno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        b_type = self.btype.get()
        if not b_type:
            messagebox.showerror(
                title="Error", message="Select a Bill Type!!")
            return
        Path("Bills/Sales/" + b_year + "/Excel"
             ).mkdir(parents=True, exist_ok=True)
        Path("Bills/Sales/" + b_year + "/PDF"
             ).mkdir(parents=True, exist_ok=True)
        if b_type == 1:
            self.generateNormalBill((b_year, b_no))
        elif b_type == 2:
            self.generateAMCBill((b_year, b_no))
        elif b_type == 3:
            self.generateHSNBill((b_year, b_no))
        elif b_type == 4:
            self.generateIGSTBill((b_year, b_no))
        else:
            self.generateAmcHsnBill((b_year, b_no))

    def generateNormalBill(self, bill_info):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select c_id, b_date FROM bill WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchone()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill does not exist!!")
                sqliteConnection.commit()
                cursor.close()
                if (sqliteConnection):
                    sqliteConnection.close()
            c_id = rows[0]
            b_date = rows[1]
            cursor.execute(
                "Select c_name, c_address, c_gst FROM client WHERE c_id=?", (c_id,))
            rows = cursor.fetchall()
            c_name, c_address, c_gst = rows[0]
            cursor.execute(
                "Select bd_id, bd_product, bd_quantity, bd_rate, bd_cgst, bd_amount FROM bill_detail WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchall()
            bill_no = bill_info[0]+'/'+str(bill_info[1])
            workbook = xlsxwriter.Workbook(
                'Bills/Sales/' + bill_info[0]+'/Excel/'+str(bill_info[1])+'-'+c_name + '.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.set_paper(9)
            worksheet.set_portrait()
            worksheet.center_horizontally()
            worksheet.set_margins(0.05, 0.05, 1.9, 0.2)
            worksheet.set_default_row(15)
            worksheet.set_header(header)
            merge_head = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 20})
            bold_11 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 11})
            bold_12 = workbook.add_format({
                'bold': 'bold',
                # 'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            bold_14 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_14_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_12_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            border = workbook.add_format({
                'border': 1})
            normal_12 = workbook.add_format({
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_header = workbook.add_format({
                'border': 1,
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_data = workbook.add_format({
                'border': 1,
                'valign': 'bottom',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 11})

            # ------ Set column Width -----
            # Sr. No.
            worksheet.set_column(0, 0, 5)
            # PARTICULARS / PRODUCT
            worksheet.set_column(1, 1, 42)
            # Quantity
            worksheet.set_column(2, 2, 5)
            # Rate
            worksheet.set_column(3, 3, 7)
            # GST
            worksheet.set_column(4, 4, 8)
            # CGST
            worksheet.set_column(5, 5, 8)
            # SGST
            worksheet.set_column(6, 6, 8)
            # Amount
            worksheet.set_column(7, 7, 10)

            # -------- Excelsheet --------
            len_entries = len(rows)
            rem = len(rows)
            pgs = 0
            total = 0
            gst = 0
            if len_entries > 40:
                pgs = int(len_entries/40)
                len_entries %= 40
            if len_entries > 20:
                pgs += 2
            else:
                pgs += 1
            len_entries = rem
            for p in range(0, pgs):
                pg = p*47
                worksheet.merge_range('A' + str(1+pg)+':H' + str(
                    1+pg), "GST No:- 27AADPP0622L1ZQ   Tel. No: +919820552008 / +919004023428   Email:aqpatanwala@hotmail.com", bold_11)
                worksheet.write(
                    'A' + str(3+pg), "To:- " + c_name, bold_14_u)
                worksheet.merge_range(
                    'A' + str(4+pg)+':H' + str(4+pg), "Address :- "+c_address, normal_12)
                worksheet.write(
                    'A' + str(5+pg), "GST:- " + c_gst, bold_12_u)

                # Bill No.
                worksheet.write_rich_string(
                    'D' + str(5+pg), bold_12, "Bill No.: ", normal_12, bill_no)

                # Date
                worksheet.write_rich_string(
                    'G' + str(5+pg), bold_12, "Date: ", normal_12, b_date)

                # Table Columns
                worksheet.write(
                    'A' + str(7+pg), "Sr.No.", table_header)
                worksheet.write(
                    'B' + str(7+pg), "PARTICULARS", table_header)
                worksheet.write(
                    'C' + str(7+pg), "QTY", table_header)
                worksheet.write(
                    'D' + str(7+pg), "RATE", table_header)
                worksheet.write(
                    'E' + str(7+pg), "GST%", table_header)
                worksheet.write(
                    'F' + str(7+pg), "CGST%", table_header)
                worksheet.write(
                    'G' + str(7+pg), "SGST%", table_header)
                worksheet.write(
                    'H' + str(7+pg), "AMOUNT", table_header)
                if rem > 40:
                    for i in range(pg+8, pg+48):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_number(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4]*2, table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        total += rows[ent][5]
                        gst += rows[ent][5]*rows[ent][4]*0.01
                        rem -= 1
                elif rem > 20:
                    for i in range(pg+8, pg+28):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_number(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4]*2, table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        total += rows[ent][5]
                        gst += rows[ent][5]*rows[ent][4]*0.01
                        rem -= 1
                else:
                    count = rem
                    for i in range(pg+8, pg+28):
                        if rem == 0:
                            break
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_number(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4]*2, table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        total += rows[ent][5]
                        gst += rows[ent][5]*rows[ent][4]*0.01
                        rem -= 1
                    if count < 20:
                        for i in range(0, 8):
                            draw_frame_border(workbook, worksheet,
                                              count+7+pg, i, 20-count, 1)
                    break

            # TOTAL
            gst = round(gst, 2)
            worksheet.merge_range('E' + str(28+pg)+':G' +
                                  str(28+pg), "TOTAL AMOUNT", table_header)
            worksheet.write_number('H' + str(28+pg), total, table_header)
            worksheet.merge_range('E' + str(29+pg)+':G' +
                                  str(29+pg), "TOTAL CGST", table_header)
            worksheet.write_number('H' + str(29+pg), gst, table_header)
            worksheet.merge_range('E' + str(30+pg)+':G' +
                                  str(30+pg), "TOTAL SGST", table_header)
            worksheet.write_number('H' + str(30+pg), gst, table_header)
            worksheet.merge_range('E' + str(31+pg)+':G' +
                                  str(31+pg), "Round off + / -", table_header)
            grand_total = total+(2*gst)
            total_roundup = int(
                Decimal(grand_total).quantize(0, ROUND_HALF_UP))
            worksheet.write_number('H' + str(31+pg), abs(
                total_roundup - grand_total), table_header)
            worksheet.merge_range('E' + str(32+pg)+':G' +
                                  str(32+pg), "GRAND TOTAL", table_header)
            worksheet.write_number(
                'H' + str(32+pg), total_roundup, table_header)

            worksheet.write('A' + str(34+pg), "Rupees:- " +
                            convertToWords(total_roundup), normal_12)
            worksheet.write(
                'A' + str(36+pg), "Note:- Goods once sold cannot be taken back.", normal_12)
            worksheet.write(
                'A' + str(37+pg), "Guarantee & Warranty applicable as per the original component suppliers terms & condition only.", normal_12)

            # Bank Details
            worksheet.write('A' + str(40+pg), "Bank details:-", normal_12)
            worksheet.write('A' + str(41+pg),
                            "Bank Name:- Bank of India", normal_12)
            worksheet.write('A' + str(42+pg),
                            "Branch :- Kalbadevi Branch", normal_12)
            worksheet.write('A' + str(43+pg),
                            "A/C No:- 002420110001459", normal_12)
            worksheet.write(
                'A' + str(44+pg), "RTGS/NEFT/IFSC Code: BKID0000024", normal_12)

            # Proprietery Signature
            worksheet.merge_range(
                'E' + str(39+pg)+':H' + str(39+pg), "FOR F.K. PATANWALA & Co.", bold_14)
            worksheet.merge_range(
                'E' + str(45+pg)+':H' + str(45+pg), "Proprietor/Authorized signatory", bold_14)

            cursor.execute(
                "Select bd_ch_no FROM bill_detail WHERE b_year=? AND b_no=? AND bd_ch_no IS NOT ''", bill_info)
            rows = cursor.fetchall()
            if len(rows):
                rows = sorted(set(rows))
                challan = ''
                for ch in rows:
                    challan += str(ch[0]) + ","
                worksheet.write_rich_string(
                    'A6', bold_12, "Challan No.: ", normal_12, challan)
            cursor.execute("Select b_po FROM bill WHERE b_year=? AND b_no=?;", bill_info)
            row = cursor.fetchone()
            if row[0]:
                worksheet.write_rich_string(
                    'F6', bold_12, "P.O. No.: ", normal_12, row[0])
            workbook.close()

            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Excel Bill created successfully!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get Bill details from table", message=error)
        except xlsxwriter.exceptions.FileCreateError as error:
            messagebox.showerror(
                title="Failed to generate Bill", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
        self.xlsxToPdf(bill_info[0], bill_info[1], c_name)

    def generateAMCBill(self, bill_info):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select c_id, b_date FROM bill WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchone()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill does not exist!!")
                sqliteConnection.commit()
                cursor.close()
                if (sqliteConnection):
                    sqliteConnection.close()
            c_id = rows[0]
            b_date = rows[1]
            cursor.execute(
                "Select c_name, c_address, c_gst FROM client WHERE c_id=?", (c_id,))
            rows = cursor.fetchall()
            c_name, c_address, c_gst = rows[0]
            cursor.execute(
                "Select bd_id,bd_amc_no, bd_product, bd_quantity, bd_rate, bd_cgst, bd_amount FROM bill_detail WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchall()
            bill_no = bill_info[0]+'/'+str(bill_info[1])
            workbook = xlsxwriter.Workbook(
                'Bills/Sales/' + bill_info[0]+'/Excel/'+str(bill_info[1])+'-'+c_name + '.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.set_paper(9)
            worksheet.set_portrait()
            worksheet.center_horizontally()
            worksheet.set_margins(0.05, 0.05, 1.9, 0.2)
            worksheet.set_default_row(15)
            worksheet.set_header(header)
            merge_head = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 20})
            bold_11 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 11})
            bold_12 = workbook.add_format({
                'bold': 'bold',
                # 'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            bold_14 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_14_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_12_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            border = workbook.add_format({
                'border': 1})
            normal_12 = workbook.add_format({
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_header = workbook.add_format({
                'border': 1,
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_data = workbook.add_format({
                'border': 1,
                'valign': 'bottom',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 11})

            # ------ Set column Width -----
            worksheet.set_column(0, 0, 5)  # Sr. No.
            worksheet.set_column(1, 1, 6)  # Amc No.
            worksheet.set_column(2, 2, 36)  # PARTICULARS / PRODUCT
            worksheet.set_column(3, 3, 5)  # Quantity
            worksheet.set_column(4, 4, 7)  # Rate
            worksheet.set_column(5, 5, 8)  # GST
            worksheet.set_column(6, 6, 8)  # CGST
            worksheet.set_column(7, 7, 8)  # SGST
            worksheet.set_column(8, 8, 10)  # Amount

            # -------- Excelsheet --------
            len_entries = len(rows)
            rem = len(rows)
            pgs = 0
            total = 0
            gst = 0
            if len_entries > 40:
                pgs = int(len_entries/40)
                len_entries %= 40
            if len_entries > 20:
                pgs += 2
            else:
                pgs += 1
            len_entries = rem
            for p in range(0, pgs):
                pg = p*47
                worksheet.merge_range('A' + str(1+pg)+':I' + str(
                    1+pg), "GST No:- 27AADPP0622L1ZQ   Tel. No: +919820552008 / +919004022828   Email:aqpatanwala@hotmail.com", bold_11)
                worksheet.write(
                    'A' + str(3+pg), "To:- " + c_name, bold_14_u)
                worksheet.merge_range(
                    'A' + str(4+pg)+':I' + str(4+pg), "Address :- "+c_address, normal_12)
                worksheet.write(
                    'A' + str(5+pg), "GST:- " + c_gst, bold_12_u)

                # Bill No.
                worksheet.write_rich_string(
                    'D' + str(5+pg), bold_12, "Bill No.: ", normal_12, bill_no)

                # Date
                worksheet.write_rich_string(
                    'H' + str(5+pg), bold_12, "Date: ", normal_12, b_date)

                # Table Columns
                worksheet.write(
                    'A' + str(7+pg), "Sr.No.", table_header)
                worksheet.write(
                    'B' + str(7+pg), "AMC", table_header)
                worksheet.write(
                    'C' + str(7+pg), "PARTICULARS", table_header)
                worksheet.write(
                    'D' + str(7+pg), "QTY", table_header)
                worksheet.write(
                    'E' + str(7+pg), "RATE", table_header)
                worksheet.write(
                    'F' + str(7+pg), "GST%", table_header)
                worksheet.write(
                    'G' + str(7+pg), "CGST%", table_header)
                worksheet.write(
                    'H' + str(7+pg), "SGST%", table_header)
                worksheet.write(
                    'I' + str(7+pg), "AMOUNT", table_header)
                if rem > 40:
                    for i in range(pg+8, pg+48):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5]*2, table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        total += rows[ent][6]
                        gst += rows[ent][6]*rows[ent][5]*0.01
                        rem -= 1
                elif rem > 20:
                    for i in range(pg+8, pg+28):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5]*2, table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        total += rows[ent][6]
                        gst += rows[ent][6]*rows[ent][5]*0.01
                        rem -= 1
                else:
                    count = rem
                    for i in range(pg+8, pg+28):
                        if rem == 0:
                            break
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5]*2, table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        total += rows[ent][6]
                        gst += rows[ent][6]*rows[ent][5]*0.01
                        rem -= 1
                    if count < 20:
                        for i in range(0, 9):
                            draw_frame_border(workbook, worksheet,
                                              count+7+pg, i, 20-count, 1)
                    break

            # TOTAL
            gst = round(gst, 2)
            worksheet.merge_range('F' + str(28+pg)+':H' +
                                  str(28+pg), "TOTAL AMOUNT", table_header)
            worksheet.write_number('I' + str(28+pg), total, table_header)
            worksheet.merge_range('F' + str(29+pg)+':H' +
                                  str(29+pg), "TOTAL CGST", table_header)
            worksheet.write_number('I' + str(29+pg), gst, table_header)
            worksheet.merge_range('F' + str(30+pg)+':H' +
                                  str(30+pg), "TOTAL SGST", table_header)
            worksheet.write_number('I' + str(30+pg), gst, table_header)
            worksheet.merge_range('F' + str(31+pg)+':H' +
                                  str(31+pg), "Round off + / -", table_header)
            grand_total = total+(2*gst)
            total_roundup = int(
                Decimal(grand_total).quantize(0, ROUND_HALF_UP))
            worksheet.write_number(
                'I' + str(31+pg), abs(total_roundup - grand_total), table_header)
            worksheet.merge_range('F' + str(32+pg)+':H' +
                                  str(32+pg), "GRAND TOTAL", table_header)
            worksheet.write_number(
                'I' + str(32+pg), total_roundup, table_header)

            worksheet.write('A' + str(34+pg), "Rupees:- " +
                            convertToWords(total_roundup), normal_12)
            worksheet.write(
                'A' + str(36+pg), "Note:- Goods once sold cannot be taken back.", normal_12)
            worksheet.write(
                'A' + str(37+pg), "Guarantee & Warranty applicable as per the original component suppliers terms & condition only.", normal_12)

            # Bank Details
            worksheet.write('A' + str(40+pg), "Bank details:-", normal_12)
            worksheet.write('A' + str(41+pg),
                            "Bank Name:- Bank of India", normal_12)
            worksheet.write('A' + str(42+pg),
                            "Branch :- Kalbadevi Branch", normal_12)
            worksheet.write('A' + str(43+pg),
                            "A/C No:- 002420110001459", normal_12)
            worksheet.write(
                'A' + str(44+pg), "RTGS/NEFT/IFSC Code: BKID0000024", normal_12)

            # Proprietery Signature
            worksheet.merge_range(
                'E' + str(39+pg)+':I' + str(39+pg), "FOR F.K. PATANWALA & Co.", bold_14)
            worksheet.merge_range(
                'E' + str(45+pg)+':I' + str(45+pg), "Proprietor/Authorized signatory", bold_14)
            cursor.execute(
                "Select bd_ch_no FROM bill_detail WHERE b_year=? AND b_no=? AND bd_ch_no IS NOT ''", bill_info)
            rows = cursor.fetchall()
            if len(rows):
                rows = sorted(set(rows))
                challan = ''
                for ch in rows:
                    challan += str(ch[0]) + ","
                worksheet.write_rich_string(
                    'A6', bold_12, "Challan No.: ", normal_12, challan)
            cursor.execute("Select b_po FROM bill WHERE b_year=? AND b_no=?;", bill_info)
            row = cursor.fetchone()
            if row[0]:
                worksheet.write_rich_string(
                    'G6', bold_12, "P.O. No.: ", normal_12, row[0])
            workbook.close()

            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Excel Bill created successfully!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get Bill details from table", message=error)
        except xlsxwriter.exceptions.FileCreateError as error:
            messagebox.showerror(
                title="Failed to generate Bill", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
        self.xlsxToPdf(bill_info[0], bill_info[1], c_name)

    def generateHSNBill(self, bill_info):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select c_id, b_date FROM bill WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchone()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill does not exist!!")
                sqliteConnection.commit()
                cursor.close()
                if (sqliteConnection):
                    sqliteConnection.close()
            c_id = rows[0]
            b_date = rows[1]
            cursor.execute(
                "Select c_name, c_address, c_gst FROM client WHERE c_id=?", (c_id,))
            rows = cursor.fetchall()
            c_name, c_address, c_gst = rows[0]
            cursor.execute(
                "Select bd_id, bd_product, bd_hsn, bd_quantity, bd_rate, bd_cgst, bd_amount FROM bill_detail WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchall()
            bill_no = bill_info[0]+'/'+str(bill_info[1])
            workbook = xlsxwriter.Workbook(
                'Bills/Sales/' + bill_info[0]+'/Excel/'+str(bill_info[1])+'-'+c_name + '.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.set_paper(9)
            worksheet.set_portrait()
            worksheet.center_horizontally()
            worksheet.set_margins(0.05, 0.05, 1.9, 0.2)
            worksheet.set_default_row(15)
            worksheet.set_header(header)
            merge_head = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 20})
            bold_11 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 11})
            bold_12 = workbook.add_format({
                'bold': 'bold',
                # 'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            bold_14 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_14_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_12_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            border = workbook.add_format({
                'border': 1})
            normal_12 = workbook.add_format({
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_header = workbook.add_format({
                'border': 1,
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_data = workbook.add_format({
                'border': 1,
                'valign': 'bottom',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 11})

            # ------ Set column Width -----
            worksheet.set_column(0, 0, 5)  # Sr. No.
            worksheet.set_column(1, 1, 36)  # PARTICULARS / PRODUCT
            worksheet.set_column(2, 2, 6)  # HSN
            worksheet.set_column(3, 3, 5)  # Quantity
            worksheet.set_column(4, 4, 7)  # Rate
            worksheet.set_column(5, 5, 8)  # GST
            worksheet.set_column(6, 6, 8)  # CGST
            worksheet.set_column(7, 7, 8)  # SGST
            worksheet.set_column(8, 8, 10)  # Amount

            # -------- Excelsheet --------
            len_entries = len(rows)
            rem = len(rows)
            pgs = 0
            total = 0
            gst = 0
            if len_entries > 40:
                pgs = int(len_entries/40)
                len_entries %= 40
            if len_entries > 20:
                pgs += 2
            else:
                pgs += 1
            len_entries = rem
            for p in range(0, pgs):
                pg = p*47
                worksheet.merge_range('A' + str(1+pg)+':I' + str(
                    1+pg), "GST No:- 27AADPP0622L1ZQ   Tel. No: +919820552008 / +919004023128   Email:aqpatanwala@hotmail.com", bold_11)
                worksheet.write(
                    'A' + str(3+pg), "To:- " + c_name, bold_14_u)
                worksheet.merge_range(
                    'A' + str(4+pg)+':I' + str(4+pg), "Address :- "+c_address, normal_12)
                worksheet.write(
                    'A' + str(5+pg), "GST:- " + c_gst, bold_12_u)

                # Bill No.
                worksheet.write_rich_string(
                    'D' + str(5+pg), bold_12, "Bill No.: ", normal_12, bill_no)

                # Date
                worksheet.write_rich_string(
                    'H' + str(5+pg), bold_12, "Date: ", normal_12, b_date)

                # Table Columns
                worksheet.write(
                    'A' + str(7+pg), "Sr.No.", table_header)
                worksheet.write(
                    'B' + str(7+pg), "PARTICULARS", table_header)
                worksheet.write(
                    'C' + str(7+pg), "HSN", table_header)
                worksheet.write(
                    'D' + str(7+pg), "QTY", table_header)
                worksheet.write(
                    'E' + str(7+pg), "RATE", table_header)
                worksheet.write(
                    'F' + str(7+pg), "GST%", table_header)
                worksheet.write(
                    'G' + str(7+pg), "CGST%", table_header)
                worksheet.write(
                    'H' + str(7+pg), "SGST%", table_header)
                worksheet.write(
                    'I' + str(7+pg), "AMOUNT", table_header)
                if rem > 40:
                    for i in range(pg+8, pg+48):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5]*2, table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        total += rows[ent][6]
                        gst += rows[ent][6]*rows[ent][5]*0.01
                        rem -= 1
                elif rem > 20:
                    for i in range(pg+8, pg+28):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5]*2, table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        total += rows[ent][6]
                        gst += rows[ent][6]*rows[ent][5]*0.01
                        rem -= 1
                else:
                    count = rem
                    for i in range(pg+8, pg+28):
                        if rem == 0:
                            break
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5]*2, table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        total += rows[ent][6]
                        gst += rows[ent][6]*rows[ent][5]*0.01
                        rem -= 1
                    if count < 20:
                        for i in range(0, 9):
                            draw_frame_border(workbook, worksheet,
                                              count+7+pg, i, 20-count, 1)
                    break

            # TOTAL
            gst = round(gst, 2)
            worksheet.merge_range('F' + str(28+pg)+':H' +
                                  str(28+pg), "TOTAL AMOUNT", table_header)
            worksheet.write_number('I' + str(28+pg), total, table_header)
            worksheet.merge_range('F' + str(29+pg)+':H' +
                                  str(29+pg), "TOTAL CGST", table_header)
            worksheet.write_number('I' + str(29+pg), gst, table_header)
            worksheet.merge_range('F' + str(30+pg)+':H' +
                                  str(30+pg), "TOTAL SGST", table_header)
            worksheet.write_number('I' + str(30+pg), gst, table_header)
            worksheet.merge_range('F' + str(31+pg)+':H' +
                                  str(31+pg), "Round off + / -", table_header)
            grand_total = total+(2*gst)
            total_roundup = int(
                Decimal(grand_total).quantize(0, ROUND_HALF_UP))
            worksheet.write_number(
                'I' + str(31+pg), abs(total_roundup - grand_total), table_header)
            worksheet.merge_range('F' + str(32+pg)+':H' +
                                  str(32+pg), "GRAND TOTAL", table_header)
            worksheet.write_number(
                'I' + str(32+pg), total_roundup, table_header)

            worksheet.write('A' + str(34+pg), "Rupees:- " +
                            convertToWords(total_roundup), normal_12)
            worksheet.write(
                'A' + str(36+pg), "Note:- Goods once sold cannot be taken back.", normal_12)
            worksheet.write(
                'A' + str(37+pg), "Guarantee & Warranty applicable as per the original component suppliers terms & condition only.", normal_12)

            # Bank Details
            worksheet.write('A' + str(40+pg), "Bank details:-", normal_12)
            worksheet.write('A' + str(41+pg),
                            "Bank Name:- Bank of India", normal_12)
            worksheet.write('A' + str(42+pg),
                            "Branch :- Kalbadevi Branch", normal_12)
            worksheet.write('A' + str(43+pg),
                            "A/C No:- 002420110001459", normal_12)
            worksheet.write(
                'A' + str(44+pg), "RTGS/NEFT/IFSC Code: BKID0000024", normal_12)

            # Proprietery Signature
            worksheet.merge_range(
                'E' + str(39+pg)+':I' + str(39+pg), "FOR F.K. PATANWALA & Co.", bold_14)
            worksheet.merge_range(
                'E' + str(45+pg)+':I' + str(45+pg), "Proprietor/Authorized signatory", bold_14)
            cursor.execute(
                "Select bd_ch_no FROM bill_detail WHERE b_year=? AND b_no=? AND bd_ch_no IS NOT ''", bill_info)
            rows = cursor.fetchall()
            if len(rows):
                rows = sorted(set(rows))
                challan = ''
                for ch in rows:
                    challan += str(ch[0]) + ","
                worksheet.write_rich_string(
                    'A6', bold_12, "Challan No.: ", normal_12, challan)
            cursor.execute("Select b_po FROM bill WHERE b_year=? AND b_no=?;", bill_info)
            row = cursor.fetchone()
            if row[0]:
                worksheet.write_rich_string(
                    'G6', bold_12, "P.O. No.: ", normal_12, row[0])
            workbook.close()

            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Excel Bill created successfully!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get Bill details from table", message=error)
        except xlsxwriter.exceptions.FileCreateError as error:
            messagebox.showerror(
                title="Failed to generate Bill", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
        self.xlsxToPdf(bill_info[0], bill_info[1], c_name)

    def generateIGSTBill(self, bill_info):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select c_id, b_date FROM bill WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchone()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill does not exist!!")
                sqliteConnection.commit()
                cursor.close()
                if (sqliteConnection):
                    sqliteConnection.close()
            c_id = rows[0]
            b_date = rows[1]
            cursor.execute(
                "Select c_name, c_address, c_gst FROM client WHERE c_id=?", (c_id,))
            rows = cursor.fetchall()
            c_name, c_address, c_gst = rows[0]
            cursor.execute(
                "Select bd_id, bd_product, bd_quantity, bd_rate, bd_igst, bd_amount FROM bill_detail WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchall()
            bill_no = bill_info[0]+'/'+str(bill_info[1])
            workbook = xlsxwriter.Workbook(
                'Bills/Sales/' + bill_info[0]+'/Excel/'+str(bill_info[1])+'-'+c_name + '.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.set_paper(9)
            worksheet.set_portrait()
            worksheet.center_horizontally()
            worksheet.set_margins(0.05, 0.05, 1.9, 0.2)
            worksheet.set_default_row(15)
            worksheet.set_header(header)
            merge_head = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 20})
            bold_11 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 11})
            bold_12 = workbook.add_format({
                'bold': 'bold',
                # 'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            bold_14 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_14_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_12_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            border = workbook.add_format({
                'border': 1})
            normal_12 = workbook.add_format({
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_header = workbook.add_format({
                'border': 1,
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_data = workbook.add_format({
                'border': 1,
                'valign': 'bottom',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 11})

            # ------ Set column Width -----
            worksheet.set_column(0, 0, 5)  # Sr. No.
            worksheet.set_column(1, 1, 50)  # PARTICULARS / PRODUCT
            worksheet.set_column(2, 2, 5)  # Quantity
            worksheet.set_column(3, 3, 7)  # Rate
            worksheet.set_column(4, 4, 8)  # GST
            worksheet.set_column(5, 5, 8)  # IGST
            worksheet.set_column(6, 6, 10)  # Amount

            # -------- Excelsheet --------
            len_entries = len(rows)
            rem = len(rows)
            pgs = 0
            total = 0
            gst = 0
            igst = 0
            if len_entries > 40:
                pgs = int(len_entries/40)
                len_entries %= 40
            if len_entries > 20:
                pgs += 2
            else:
                pgs += 1
            len_entries = rem
            for p in range(0, pgs):
                pg = p*47
                worksheet.merge_range('A' + str(1+pg)+':G' + str(
                    1+pg), "GST No:- 27AADPP0622L1ZQ   Tel. No: +919820552008 / +919004020428   Email:aqpatanwala@hotmail.com", bold_11)
                worksheet.write(
                    'A' + str(3+pg), "To:- " + c_name, bold_14_u)
                worksheet.merge_range(
                    'A' + str(4+pg)+':G' + str(4+pg), "Address :- "+c_address, normal_12)
                # Bill No.
                worksheet.write_rich_string(
                    'A' + str(5+pg), bold_12_u, "GST:- " + c_gst, bold_12, "              Bill No.: ", normal_12, bill_no)

                # Date
                worksheet.write_rich_string(
                    'F' + str(5+pg), bold_12, "Date: ", normal_12, b_date)

                # Table Columns
                worksheet.write(
                    'A' + str(7+pg), "Sr.No.", table_header)
                worksheet.write(
                    'B' + str(7+pg), "PARTICULARS", table_header)
                worksheet.write(
                    'C' + str(7+pg), "QTY", table_header)
                worksheet.write(
                    'D' + str(7+pg), "RATE", table_header)
                worksheet.write(
                    'E' + str(7+pg), "GST%", table_header)
                worksheet.write(
                    'F' + str(7+pg), "IGST%", table_header)
                worksheet.write(
                    'G' + str(7+pg), "AMOUNT", table_header)
                if rem > 40:
                    for i in range(pg+8, pg+48):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_number(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        total += rows[ent][5]
                        igst += rows[ent][5]*rows[ent][4]*0.01
                        rem -= 1
                elif rem > 20:
                    for i in range(pg+8, pg+28):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_number(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        total += rows[ent][5]
                        igst += rows[ent][5]*rows[ent][4]*0.01
                        rem -= 1
                else:
                    count = rem
                    for i in range(pg+8, pg+28):
                        if rem == 0:
                            break
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_number(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_number(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][5], table_data)
                        total += rows[ent][5]
                        igst += rows[ent][5]*rows[ent][4]*0.01
                        rem -= 1
                    if count < 20:
                        for i in range(0, 10):
                            draw_frame_border(workbook, worksheet,
                                              count+7+pg, i, 20-count, 1)
                    break

            # TOTAL
            igst = round(igst, 2)
            worksheet.merge_range('D' + str(28+pg)+':F' +
                                  str(28+pg), "TOTAL AMOUNT", table_header)
            worksheet.write_number('G' + str(28+pg), total, table_header)
            worksheet.merge_range('D' + str(29+pg)+':F' +
                                  str(29+pg), "TOTAL IGST", table_header)
            worksheet.write_number('G' + str(29+pg), igst, table_header)
            worksheet.merge_range('D' + str(30+pg)+':F' +
                                  str(30+pg), "Round off + / -", table_header)
            grand_total = total + igst
            total_roundup = int(
                Decimal(grand_total).quantize(0, ROUND_HALF_UP))
            worksheet.write_number(
                'G' + str(30+pg), abs(total_roundup - grand_total), table_header)
            worksheet.merge_range('D' + str(31+pg)+':F' +
                                  str(31+pg), "GRAND TOTAL", table_header)
            worksheet.write_number(
                'G' + str(31+pg), total_roundup, table_header)

            worksheet.write('A' + str(33+pg), "Rupees:- " +
                            convertToWords(total_roundup), normal_12)
            worksheet.write(
                'A' + str(36+pg), "Note:- Goods once sold cannot be taken back.", normal_12)
            worksheet.write(
                'A' + str(37+pg), "Guarantee & Warranty applicable as per the original component suppliers terms & condition only.", normal_12)

            # Bank Details
            worksheet.write('A' + str(40+pg), "Bank details:-", normal_12)
            worksheet.write('A' + str(41+pg),
                            "Bank Name:- Bank of India", normal_12)
            worksheet.write('A' + str(42+pg),
                            "Branch :- Kalbadevi Branch", normal_12)
            worksheet.write('A' + str(43+pg),
                            "A/C No:- 002420110001459", normal_12)
            worksheet.write(
                'A' + str(44+pg), "RTGS/NEFT/IFSC Code: BKID0000024", normal_12)

            # Proprietery Signature
            worksheet.merge_range(
                'D' + str(39+pg)+':G' + str(39+pg), "FOR F.K. PATANWALA & Co.", bold_14)
            worksheet.merge_range(
                'D' + str(45+pg)+':G' + str(45+pg), "Proprietor/Authorized signatory", bold_14)
            cursor.execute(
                "Select bd_ch_no FROM bill_detail WHERE b_year=? AND b_no=? AND bd_ch_no IS NOT ''", bill_info)
            rows = cursor.fetchall()
            if len(rows):
                rows = sorted(set(rows))
                challan = ''
                for ch in rows:
                    challan += str(ch[0]) + ","
                worksheet.write_rich_string(
                    'A6', bold_12, "Challan No.: ", normal_12, challan)
            row = cursor.fetchone()
            if row[0]:
                worksheet.write_rich_string(
                    'E6', bold_12, "P.O. No.: ", normal_12, row[0])
            workbook.close()

            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Excel Bill created successfully!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get Bill details from table", message=error)
        except xlsxwriter.exceptions.FileCreateError as error:
            messagebox.showerror(
                title="Failed to generate Bill", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
        self.xlsxToPdf(bill_info[0], bill_info[1], c_name)

    def generateAmcHsnBill(self, bill_info):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select c_id, b_date FROM bill WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchone()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill does not exist!!")
                sqliteConnection.commit()
                cursor.close()
                if (sqliteConnection):
                    sqliteConnection.close()
            c_id = rows[0]
            b_date = rows[1]
            cursor.execute(
                "Select c_name, c_address, c_gst FROM client WHERE c_id=?", (c_id,))
            rows = cursor.fetchall()
            c_name, c_address, c_gst = rows[0]
            cursor.execute(
                "Select bd_id,bd_amc_no, bd_product, bd_hsn, bd_quantity, bd_rate, bd_cgst, bd_amount FROM bill_detail WHERE b_year=? AND b_no=?", bill_info)
            rows = cursor.fetchall()
            bill_no = bill_info[0]+'/'+str(bill_info[1])
            workbook = xlsxwriter.Workbook(
                'Bills/Sales/' + bill_info[0]+'/Excel/'+str(bill_info[1])+'-'+c_name + '.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.set_paper(9)
            worksheet.set_portrait()
            worksheet.center_horizontally()
            worksheet.set_margins(0.05, 0.05, 1.9, 0.2)
            worksheet.set_default_row(15)
            worksheet.set_header(header)
            merge_head = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 20})
            bold_11 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 11})
            bold_12 = workbook.add_format({
                'bold': 'bold',
                # 'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            bold_14 = workbook.add_format({
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_14_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 14})
            bold_12_u = workbook.add_format({
                'bold': 'bold',
                'underline': 'underline',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            border = workbook.add_format({
                'border': 1})
            normal_12 = workbook.add_format({
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_header = workbook.add_format({
                'border': 1,
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_data = workbook.add_format({
                'border': 1,
                'valign': 'bottom',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 11})

            # ------ Set column Width -----
            worksheet.set_column(0, 0, 5)  # Sr. No.
            worksheet.set_column(1, 1, 6)  # Amc No.
            worksheet.set_column(2, 2, 31)  # PARTICULARS / PRODUCT
            worksheet.set_column(3, 3, 5)  # HSN
            worksheet.set_column(4, 4, 5)  # Quantity
            worksheet.set_column(5, 5, 7)  # Rate
            worksheet.set_column(6, 6, 8)  # GST
            worksheet.set_column(7, 7, 8)  # CGST
            worksheet.set_column(8, 8, 8)  # SGST
            worksheet.set_column(9, 9, 10)  # Amount
            # -------- Excelsheet --------
            len_entries = len(rows)
            rem = len(rows)
            pgs = 0
            total = 0
            gst = 0
            if len_entries > 40:
                pgs = int(len_entries/40)
                len_entries %= 40
            if len_entries > 20:
                pgs += 2
            else:
                pgs += 1
            len_entries = rem
            for p in range(0, pgs):
                pg = p*47
                worksheet.merge_range('A' + str(1+pg)+':J' + str(
                    1+pg), "GST No:- 27AADPP0622L1ZQ   Tel. No: +919820552008 / +919004023128   Email:aqpatanwala@hotmail.com", bold_11)
                worksheet.write(
                    'A' + str(3+pg), "To:- " + c_name, bold_14_u)
                worksheet.merge_range(
                    'A' + str(4+pg)+':J' + str(4+pg), "Address :- "+c_address, normal_12)
                worksheet.write(
                    'A' + str(5+pg), "GST:- " + c_gst, bold_12_u)

                # Bill No.
                worksheet.write_rich_string(
                    'D' + str(5+pg), bold_12, "Bill No.: ", normal_12, bill_no)

                # Date
                worksheet.write_rich_string(
                    'H' + str(5+pg), bold_12, "Date: ", normal_12, b_date)

                # Table Columns
                worksheet.write(
                    'A' + str(7+pg), "Sr.No.", table_header)
                worksheet.write(
                    'B' + str(7+pg), "AMC", table_header)
                worksheet.write(
                    'C' + str(7+pg), "PARTICULARS", table_header)
                worksheet.write(
                    'D' + str(7+pg), "HSN", table_header)
                worksheet.write(
                    'E' + str(7+pg), "QTY", table_header)
                worksheet.write(
                    'F' + str(7+pg), "RATE", table_header)
                worksheet.write(
                    'G' + str(7+pg), "GST%", table_header)
                worksheet.write(
                    'H' + str(7+pg), "CGST%", table_header)
                worksheet.write(
                    'I' + str(7+pg), "SGST%", table_header)
                worksheet.write(
                    'J' + str(7+pg), "AMOUNT", table_header)
                if rem > 40:
                    for i in range(pg+8, pg+48):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_string(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][6]*2, table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][6], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        worksheet.write_number(
                            'J' + str(i), rows[ent][7], table_data)
                        total += rows[ent][7]
                        gst += rows[ent][7]*rows[ent][6]*0.01
                        rem -= 1
                elif rem > 20:
                    for i in range(pg+8, pg+28):
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_string(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][6]*2, table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][6], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        worksheet.write_number(
                            'J' + str(i), rows[ent][7], table_data)
                        total += rows[ent][7]
                        gst += rows[ent][7]*rows[ent][6]*0.01
                        rem -= 1
                else:
                    count = rem
                    for i in range(pg+8, pg+28):
                        if rem == 0:
                            break
                        ent = len_entries-rem
                        worksheet.write_number(
                            'A' + str(i), rows[ent][0], table_data)
                        worksheet.write_string(
                            'B' + str(i), rows[ent][1], table_data)
                        worksheet.write_string(
                            'C' + str(i), rows[ent][2], table_data)
                        worksheet.write_string(
                            'D' + str(i), rows[ent][3], table_data)
                        worksheet.write_number(
                            'E' + str(i), rows[ent][4], table_data)
                        worksheet.write_number(
                            'F' + str(i), rows[ent][5], table_data)
                        worksheet.write_number(
                            'G' + str(i), rows[ent][6]*2, table_data)
                        worksheet.write_number(
                            'H' + str(i), rows[ent][6], table_data)
                        worksheet.write_number(
                            'I' + str(i), rows[ent][6], table_data)
                        worksheet.write_number(
                            'J' + str(i), rows[ent][7], table_data)
                        total += rows[ent][7]
                        gst += rows[ent][7]*rows[ent][6]*0.01
                        rem -= 1
                    if count < 20:
                        for i in range(0, 10):
                            draw_frame_border(workbook, worksheet,
                                              count+7+pg, i, 20-count, 1)
                    break

            # TOTAL
            gst = round(gst, 2)
            worksheet.merge_range('G' + str(28+pg)+':I' +
                                  str(28+pg), "TOTAL AMOUNT", table_header)
            worksheet.write_number('J' + str(28+pg), total, table_header)
            worksheet.merge_range('G' + str(29+pg)+':I' +
                                  str(29+pg), "TOTAL CGST", table_header)
            worksheet.write_number('J' + str(29+pg), gst, table_header)
            worksheet.merge_range('G' + str(30+pg)+':I' +
                                  str(30+pg), "TOTAL SGST", table_header)
            worksheet.write_number('J' + str(30+pg), gst, table_header)
            worksheet.merge_range('G' + str(31+pg)+':I' +
                                  str(31+pg), "Round off + / -", table_header)
            grand_total = total+(2*gst)
            total_roundup = int(
                Decimal(grand_total).quantize(0, ROUND_HALF_UP))
            worksheet.write_number(
                'J' + str(31+pg), abs(total_roundup - grand_total), table_header)
            worksheet.merge_range('G' + str(32+pg)+':I' +
                                  str(32+pg), "GRAND TOTAL", table_header)
            worksheet.write_number(
                'J' + str(32+pg), total_roundup, table_header)

            worksheet.write('A' + str(34+pg), "Rupees:- " +
                            convertToWords(total_roundup), normal_12)
            worksheet.write(
                'A' + str(36+pg), "Note:- Goods once sold cannot be taken back.", normal_12)
            worksheet.write(
                'A' + str(37+pg), "Guarantee & Warranty applicable as per the original component suppliers terms & condition only.", normal_12)

            # Bank Details
            worksheet.write('A' + str(40+pg), "Bank details:-", normal_12)
            worksheet.write('A' + str(41+pg),
                            "Bank Name:- Bank of India", normal_12)
            worksheet.write('A' + str(42+pg),
                            "Branch :- Kalbadevi Branch", normal_12)
            worksheet.write('A' + str(43+pg),
                            "A/C No:- 002420110001459", normal_12)
            worksheet.write(
                'A' + str(44+pg), "RTGS/NEFT/IFSC Code: BKID0000024", normal_12)

            # Proprietery Signature
            worksheet.merge_range(
                'F' + str(39+pg)+':J' + str(39+pg), "FOR F.K. PATANWALA & Co.", bold_14)
            worksheet.merge_range(
                'F' + str(45+pg)+':J' + str(45+pg), "Proprietor/Authorized signatory", bold_14)
            cursor.execute(
                "Select bd_ch_no FROM bill_detail WHERE b_year=? AND b_no=? AND bd_ch_no IS NOT ''", bill_info)
            rows = cursor.fetchall()
            if len(rows):
                rows = sorted(set(rows))
                challan = ''
                for ch in rows:
                    challan += str(ch[0]) + ","
                worksheet.write_rich_string(
                    'A6', bold_12, "Challan No.: ", normal_12, challan)
            row = cursor.fetchone()
            if row[0]:
                worksheet.write_rich_string(
                    'H6', bold_12, "P.O. No.: ", normal_12, row[0])
            workbook.close()

            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Excel Bill created successfully!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get Bill details from table", message=error)
        except xlsxwriter.exceptions.FileCreateError as error:
            messagebox.showerror(
                title="Failed to generate Bill", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
        self.xlsxToPdf(bill_info[0], bill_info[1], c_name)

    # Covert excel to pdf
    def xlsxToPdf(self, year, num, name):
        cwd = os.getcwd()
        input_file = cwd + r'\Bills\Sales'+'\\' + year + \
            r'\Excel'+'\\'+str(num)+'-'+name+'.xlsx'
        # give your file name with valid path
        output_file = cwd + r'\Bills\Sales'+'\\' + \
            year + r'\PDF'+'\\'+str(num)+'-'+name+'.pdf'
        # give valid output file name and path
        app = client.Dispatch("Excel.Application")
        app.Interactive = False
        app.Visible = False
        Workbook = app.Workbooks.Open(input_file)
        try:
            Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
            messagebox.showinfo(title="Successful",
                                message="PDF Bill created successfully!!")
        except Exception as e:
            messagebox.showerror(
                title="Failed to convert in PDF format.", message=str(e))
        finally:
            Workbook.Close()
            app.Quit()
            del app

    # Print Bill
    def printBill(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Financial Year Field cannot be empty!!")
            return
        if b_year not in years:
            messagebox.showerror(
                title="Error", message="Select Blling Year from Dropdown List!!")
            return
        b_no = self.billno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "Select c_id FROM bill WHERE b_year=? AND b_no=?", (b_year, b_no,))
            rows = cursor.fetchone()
            c_id = rows[0]
            if not c_id:
                messagebox.showerror(
                    title="Error", message="Bill No. does not exit!!")
            cursor.execute(
                "Select c_name FROM client WHERE c_id=?", (c_id,))
            rows = cursor.fetchone()
            c_name = rows[0]

            sqliteConnection.commit()
            getCientList()
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get Bill details from table", message=error)

        finally:
            if (sqliteConnection):
                sqliteConnection.close()

        self.printPDF(b_year, b_no, c_name)

    def printPDF(self, year, num, name):
        cwd = os.getcwd()
        pdf_file = cwd + r'\Bills\Sales'+'\\' + \
            year + r'\PDF'+'\\'+str(num)+'-'+name+'.pdf'
        try:
            win32api.ShellExecute(0, "print", pdf_file, None,  ".",  0)
        except win32api.error as e:
            messagebox.showerror(title="Failed to Print Bill", message=e)


# ---------- PURCHASES FRAMES --------


class AddPurchaseBill(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Purchases Options ---------
        self.purchaseF = LabelFrame(self, text="Purchases", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.purchaseF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.purchaseF, text="Bill Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.01, relwidth=0.9, relheight=0.07)

        # -------- Purchaser's Details Frame ---------
        self.purchaseDF = LabelFrame(self.purchaseF, text="Purchaser's Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"))
        self.purchaseDF.place(relx=0.01, rely=0.08,
                              relwidth=0.98, relheight=0.15)

        # Purchaser's Name
        Label(self.purchaseDF, text="Purchaser's Name:", font=(
            "times new roman", 18, "bold")).place(relx=0.01, rely=0.2, relwidth=0.16, relheight=0.6)
        self.pname = StringVar()
        self.pnameC = ttk.Combobox(self.purchaseDF, font=(
            "arial", 18, "bold"), textvariable=self.pname, postcommand=self.updatePurchaserList)
        self.pnameC.place(relx=0.18, rely=0.2, relwidth=0.32, relheight=0.6)

        # Purchaser's GST
        Label(self.purchaseDF, text="GST No.:", font=(
            "times new roman", 18, "bold")).place(relx=0.5, rely=0.2, relwidth=0.1, relheight=0.6)
        self.gst = Entry(self.purchaseDF, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"))
        self.gst.place(relx=0.6, rely=0.2, relwidth=0.2, relheight=0.6)

        # Add Purchaser Button
        Button(self.purchaseDF, text="Add Purchaser", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=self.addPurchaser).place(relx=0.81, rely=0.15, relwidth=0.18, relheight=0.7)

        # -------- Purchase Bill Details Frame ---------
        self.purchaseBF = LabelFrame(self.purchaseF, text="Purchase Bill Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"))
        self.purchaseBF.place(relx=0.01, rely=0.23,
                              relwidth=0.98, relheight=0.75)

        # Bill No.
        Label(self.purchaseBF, text="Bill No.:", font=(
            "times new roman", 18, "bold")).place(relx=0.01, rely=0.01, relwidth=0.1, relheight=0.06)
        self.bno = StringVar()
        self.bnoE = Entry(self.purchaseBF, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"), textvariable=self.bno)
        self.bnoE.place(relx=0.11, rely=0.01, relwidth=0.15, relheight=0.07)

        # Billing Day
        Label(self.purchaseBF, text="Billing Day:", font=(
            "times new roman", 18, "bold"),).place(relx=0.26, rely=0.01, relwidth=0.15, relheight=0.06)
        self.bday = IntVar()
        self.bdayE = Entry(self.purchaseBF, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"), textvariable=self.bday)
        self.bdayE.place(relx=0.41, rely=0.01, relwidth=0.05, relheight=0.07)

        # Billing Month
        Label(self.purchaseBF, text="Billing Month:", font=(
            "times new roman", 18, "bold")).place(relx=0.48, rely=0.01, relwidth=0.15, relheight=0.06)
        self.bmonth = IntVar()
        self.bmonthE = Entry(self.purchaseBF, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"), textvariable=self.bmonth)
        self.bmonthE.place(relx=0.64, rely=0.01, relwidth=0.05, relheight=0.07)

        # Billing Year
        Label(self.purchaseBF, text="Billing Year:", font=(
            "times new roman", 18, "bold"), pady=10).place(relx=0.71, rely=0.01, relwidth=0.15, relheight=0.06)
        self.byear = IntVar()
        self.byearE = Entry(self.purchaseBF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"), textvariable=self.byear)
        self.byearE.place(relx=0.86, rely=0.01, relwidth=0.1, relheight=0.07)

        # Taxes Frame
        self.taxF = LabelFrame(self.purchaseBF, text="Taxes", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.taxF.place(relx=0.01, rely=0.09, relwidth=0.7, relheight=0.89)

        # Columns
        Label(self.taxF, text="Taxable Amount", font=(
            "times new roman", 20, "bold")).place(relx=0.1, rely=0.01, relwidth=0.24, relheight=0.1)
        Label(self.taxF, text="CGST", font=(
            "times new roman", 20, "bold")).place(relx=0.36, rely=0.01, relwidth=0.2, relheight=0.1)
        Label(self.taxF, text="SGST", font=(
            "times new roman", 20, "bold")).place(relx=0.58, rely=0.01, relwidth=0.2, relheight=0.1)
        Label(self.taxF, text="IGST", font=(
            "times new roman", 20, "bold")).place(relx=0.8, rely=0.01, relwidth=0.2, relheight=0.1)

        # Rows
        Label(self.taxF, text="5%", font=(
            "times new roman", 20, "bold")).place(relx=0, rely=0.15, relwidth=0.1, relheight=0.2)
        Label(self.taxF, text="12%", font=(
            "times new roman", 20, "bold")).place(relx=0, rely=0.35, relwidth=0.1, relheight=0.2)
        Label(self.taxF, text="18%", font=(
            "times new roman", 20, "bold")).place(relx=0, rely=0.55, relwidth=0.1, relheight=0.2)
        Label(self.taxF, text="28%", font=(
            "times new roman", 20, "bold")).place(relx=0, rely=0.75, relwidth=0.1, relheight=0.2)

        # GST 5%
        self.bamt_5 = DoubleVar()
        self.bamt_5E = Entry(self.taxF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 20, "bold"), textvariable=self.bamt_5)
        self.bamt_5E.place(relx=0.12, rely=0.2, relwidth=0.2, relheight=0.1)
        self.bgst_5 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_5).place(relx=0.38, rely=0.2, relwidth=0.16, relheight=0.1)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_5).place(relx=0.6, rely=0.2, relwidth=0.16, relheight=0.1)
        self.bigst_5 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_5).place(relx=0.82, rely=0.2, relwidth=0.16, relheight=0.1)
        self.bamt_5E.bind('<Return>', self.updateGST_5)

        # GST 12%
        self.bamt_12 = DoubleVar()
        self.bamt_12E = Entry(self.taxF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 20, "bold"), textvariable=self.bamt_12)
        self.bamt_12E.place(relx=0.12, rely=0.4, relwidth=0.2, relheight=0.1)
        self.bgst_12 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_12).place(relx=0.38, rely=0.4, relwidth=0.16, relheight=0.1)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_12).place(relx=0.6, rely=0.4, relwidth=0.16, relheight=0.1)
        self.bigst_12 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_12).place(relx=0.82, rely=0.4, relwidth=0.16, relheight=0.1)
        self.bamt_12E.bind('<Return>', self.updateGST_12)

        # GST 18%
        self.bamt_18 = DoubleVar()
        self.bamt_18E = Entry(self.taxF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 20, "bold"), textvariable=self.bamt_18)
        self.bamt_18E.place(relx=0.12, rely=0.6, relwidth=0.2, relheight=0.1)
        self.bgst_18 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_18).place(relx=0.38, rely=0.6, relwidth=0.16, relheight=0.1)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_18).place(relx=0.6, rely=0.6, relwidth=0.16, relheight=0.1)
        self.bigst_18 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_18).place(relx=0.82, rely=0.6, relwidth=0.16, relheight=0.1)
        self.bamt_18E.bind('<Return>', self.updateGST_18)

        # GST 28%
        self.bamt_28 = DoubleVar()
        self.bamt_28E = Entry(self.taxF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 20, "bold"), textvariable=self.bamt_28)
        self.bamt_28E.place(relx=0.12, rely=0.8, relwidth=0.2, relheight=0.1)
        self.bgst_28 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_28).place(relx=0.38, rely=0.8, relwidth=0.16, relheight=0.1)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_28).place(relx=0.6, rely=0.8, relwidth=0.16, relheight=0.1)
        self.bigst_28 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_28).place(relx=0.82, rely=0.8, relwidth=0.16, relheight=0.1)
        self.bamt_28E.bind('<Return>', self.updateGST_28)

        # Total Amount
        Label(self.purchaseBF, text="Total Amount:", font=(
            "times new roman", 18, "bold")).place(relx=0.71, rely=0.1, relwidth=0.27, relheight=0.07)
        self.btamt = IntVar()
        self.btamtE = Label(self.purchaseBF, width=15, bd=3, relief=GROOVE, fg=navy, font=(
            "arial", 18, "bold"), textvariable=self.btamt)
        self.btamtE.place(relx=0.72, rely=0.17, relwidth=0.27, relheight=0.1)

        # Status
        Label(self.purchaseBF, text="Status:", font=(
            "times new roman", 18, "bold")).place(relx=0.71, rely=0.3, relwidth=0.27, relheight=0.07)
        self.bstatus = Entry(self.purchaseBF, width=15, bd=3, relief=GROOVE, fg=navy, font=(
            "arial", 18, "bold"))
        self.bstatus.place(relx=0.72, rely=0.37, relwidth=0.27, relheight=0.1)

        # Total Amount Button
        Button(self.purchaseBF, text="Total Amount", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=self.totalBill).place(relx=0.72, rely=0.54, relwidth=0.27, relheight=0.1)

        # Add Bill Button
        Button(self.purchaseBF, text="Add Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=self.createBill).place(relx=0.72, rely=0.71, relwidth=0.27, relheight=0.1)

        # Home Button
        add_client_btn = Button(self.purchaseBF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.72, rely=0.88, relwidth=0.27, relheight=0.1)

    def addPurchaser(self):
        getPurchaserList()
        name = self.pname.get()
        if not name:
            messagebox.showerror(
                title="Error", message="Name Field cannot be empty!!")
            return
        if name in purchasers:
            messagebox.showerror(
                title="Error", message="Client's Name already Exist!!")
            return
        gst = self.gst.get()
        if not gst:
            messagebox.showerror(
                title="Error", message="GST No. Field cannot be empty!!")
            return
        self.insertPurchaser((name, gst,))

    def insertPurchaser(self, constraints):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                """INSERT INTO purchaser (p_name, p_gst) VALUES (?,?);""", constraints)
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Purchaser added successfully!!")
            cursor.close()
            getPurchaserList()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into purchaser table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def createBill(self):
        name = self.pname.get()
        if not name:
            messagebox.showerror(
                title="Error", message="Name Field cannot be empty!!")
            return
        if name not in purchasers:
            messagebox.showerror(
                title="Error", message="Add Client to the database First or Select from Dropdown List!!")
            return
        bill_no = self.bno.get()
        if not name:
            messagebox.showerror(
                title="Error", message="Bill No. Field cannot be empty!!")
            return
        day = self.bday.get()
        if not day:
            messagebox.showerror(
                title="Error", message="Enter a Day!!")
            return
        if not (0 < day < 32):
            messagebox.showerror(
                title="Error", message="Incorrect Day!!")
            return
        month = self.bmonth.get()
        if not month:
            messagebox.showerror(
                title="Error", message="Enter a Month!!")
            return
        if not (0 < month < 13):
            messagebox.showerror(
                title="Error", message="Incorrect Month!!")
            return
        year = int(self.byear.get())
        if not year:
            messagebox.showerror(
                title="Error", message="Enter a Year!!")
            return
        if not (2019 < year < 2030):
            messagebox.showerror(
                title="Error", message="Incorrect Year!!")
            return
        bgst_5 = self.bgst_5.get()
        bgst_12 = self.bgst_12.get()
        bgst_18 = self.bgst_18.get()
        bgst_28 = self.bgst_28.get()
        bigst_5 = self.bigst_5.get()
        bigst_12 = self.bigst_12.get()
        bigst_18 = self.bigst_18.get()
        bigst_28 = self.bigst_28.get()
        tax_amt = self.bamt_5.get()+self.bamt_12.get() + \
            self.bamt_18.get()+self.bamt_28.get()
        total_amt = self.btamt.get()
        if not tax_amt:
            messagebox.showerror(
                title="Error", message="Enter Taxable Amount!!")
            return
        status = self.bstatus.get()
        if not status:
            status = "Pending"
        constraints = [year, month, day, tax_amt, bgst_5, bigst_5,
                       bgst_12, bigst_12, bgst_18, bigst_18, bgst_28, bigst_28, total_amt, bill_no, status]
        self.insertBill(name, constraints)

    def insertBill(self, name, constraints):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                """Select MAX(pb_no) FROM purchase_bill WHERE pb_year=?;""", (constraints[0],))
            val = cursor.fetchall()
            if val[0][0] == None:
                b_no = 1
            else:
                b_no = val[0][0]+1
            cursor.execute(
                """Select p_id FROM purchaser WHERE p_name=?;""", (name,))
            val = cursor.fetchall()
            p_id = val[0][0]
            constraints.insert(0, p_id)
            constraints.insert(0, b_no)
            cursor.execute(
                """INSERT INTO purchase_bill (pb_no, p_id, pb_year, pb_month, pb_day, pb_tax_amt, pb_gst_5, pb_igst_5, pb_gst_12, pb_igst_12, pb_gst_18, pb_igst_18, pb_gst_28, pb_igst_28, pb_total_amt, pb_bill_no, pb_status)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);""", tuple(constraints))
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Bill added successfully with Bill Sr. No:"+str(b_no)+"!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into purchase_bill table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def updatePurchaserList(self):
        getPurchaserList()
        self.pnameC['values'] = purchasers

    def updateGST_5(self, event):
        self.bgst_5.set(self.bamt_5.get()*0.025)

    def updateGST_12(self, event):
        self.bgst_12.set(self.bamt_12.get()*0.06)

    def updateGST_18(self, event):
        self.bgst_18.set(self.bamt_18.get()*0.09)

    def updateGST_28(self, event):
        self.bgst_28.set(self.bamt_28.get()*0.14)

    def totalBill(self):
        self.btamt.set(round(self.bamt_5.get()+self.bamt_12.get()+self.bamt_18.get()+self.bamt_28.get()+2*(self.bgst_5.get()+self.bgst_12.get() +
                                                                                                           self.bgst_12.get()+self.bgst_28.get())+self.bigst_5.get()+self.bigst_12.get()+self.bigst_18.get()+self.bigst_28.get()))


class UpdatePurchaseStatus(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 36, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Purchases Options ---------
        self.purchaseF = LabelFrame(self, text="Purchases", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 28, "bold"))
        self.purchaseF.place(relx=0, rely=0.1, relwidth=1, relheight=0.9)

        Label(self.purchaseF, text="Check & Edit Bill Payment", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold")).place(relx=0.05, rely=0.01, relwidth=0.9, relheight=0.07)

        # --------- Check Bill Status Frame ------------
        self.SelectBillF = LabelFrame(self.purchaseF, text="  Search & Generate Excel Bill   ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"))
        self.SelectBillF.place(relx=0.01, rely=0.09,
                               relwidth=0.48, relheight=0.3)

        # Billing Year
        Label(self.SelectBillF, text="Bill Year:", font=(
            "times new roman", 18, "bold")).place(relx=0.01, rely=0.01, relwidth=0.24, relheight=0.2)
        self.byear = Entry(self.SelectBillF, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"))
        self.byear.place(relx=0.25, rely=0.02, relwidth=0.24, relheight=0.2)

        # Billing Month
        Label(self.SelectBillF, text="Bill Month:", font=(
            "times new roman", 18, "bold")).place(relx=0.51, rely=0.01, relwidth=0.24, relheight=0.2)
        self.bmonth = Entry(self.SelectBillF, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"))
        self.bmonth.place(relx=0.75, rely=0.02, relwidth=0.2, relheight=0.2)

        # Purchaser's Name
        Label(self.SelectBillF, text="Purchaser's Name:", font=(
            "times new roman", 18, "bold")).place(relx=0.01, rely=0.31, relwidth=0.44, relheight=0.2)
        self.pname = StringVar()
        self.pnameC = ttk.Combobox(self.SelectBillF, width=28, font=(
            "arial", 18, "bold"), textvariable=self.pname, postcommand=self.updatePurchaserList)
        self.pnameC.place(relx=0.41, rely=0.31, relwidth=0.54, relheight=0.2)

        # Search Bills Button
        Button(self.SelectBillF, text="Search Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=self.searchBill).place(relx=0.1, rely=0.61, relwidth=0.35, relheight=0.3)

        # Generate Excel Button
        Button(self.SelectBillF, text="Add to Excel", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=25, pady=20, font=(
            "arial", 18, "bold"), command=self.generatePurchaseBill).place(relx=0.55, rely=0.61, relwidth=0.35, relheight=0.3)

        # --------- Edit Bill Status Frame ------------
        self.editStatusF = LabelFrame(self.purchaseF, text="  Update Bill Payment  ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"))
        self.editStatusF.place(relx=0.51, rely=0.09,
                               relwidth=0.48, relheight=0.3)

        # Bill No.
        Label(self.editStatusF, text="Billing No:", font=(
            "times new roman", 18, "bold")).place(relx=0.01, rely=0.01, relwidth=0.36, relheight=0.3)
        self.bno = StringVar()
        self.bnoC = Entry(self.editStatusF, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"), textvariable=self.bno)
        self.bnoC.place(relx=0.38, rely=0.01, relwidth=0.38, relheight=0.2)

        # Payment Status
        Label(self.editStatusF, text="Payment Status:", font=(
            "times new roman", 18, "bold")).place(relx=0.01, rely=0.31, relwidth=0.36, relheight=0.3)
        self.estatus = Entry(self.editStatusF, width=30, bd=3, relief=GROOVE, font=(
            "arial", 18, "bold"))
        self.estatus.place(relx=0.38, rely=0.31, relwidth=0.38, relheight=0.2)

        # Current Bill Details Button
        Button(self.editStatusF, text="Update Status", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=self.editStatus).place(relx=0.1, rely=0.61, relwidth=0.35, relheight=0.3)

        # Home Button
        Button(self.editStatusF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(relx=0.55, rely=0.61, relwidth=0.35, relheight=0.3)

        # --------- View Bill Detials Frame ------------
        self.viewBillF = LabelFrame(self.purchaseF, text="View Bill Detials", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"))
        self.viewBillF.place(relx=0.01, rely=0.4,
                             relwidth=0.98, relheight=0.59)
        self.displayBillF = Frame(self.viewBillF)
        self.displayBillF.place(relx=0.01, rely=0.05,
                                relwidth=0.98, relheight=0.9)
        self.displayText = scrolledtext.ScrolledText(
            self.displayBillF, height=280, font=("Courier",
                                                 12, "bold"), padx=10, pady=10)
        self.displayText.insert(INSERT,
                                "\n\n\n\n\n\n\t\tSelect Billing Year or Purchaser's Name to see the entries!! ")
        self.displayText.configure(state='disabled')
        self.displayText.pack(side="left", fill="both", expand=True)
        Label(self.viewBillF, text="Bill No.\tYear\t\t            Purchaser's Name\t\t\t Amount\t              Payment", font=(
            "times new roman", 15, "bold")).place(relx=0.02, rely=0.01)

    def searchBill(self):
        constraints = []
        p_name = self.pname.get()
        if p_name:
            if p_name not in purchasers:
                messagebox.showerror(
                    title="Error", message="Select Purchaser's Name from Dropdown List!!")
                return
            constraints.append("p_name")
            constraints.append(p_name)
        b_year = self.byear.get()
        if b_year:
            constraints.append(b_year)
            b_month = self.bmonth.get()
            if b_month:
                constraints.append(b_month)
        if not len(constraints):
            messagebox.showerror(
                title="Error", message="Select atleast 1 critera to search!!")
            return
        self.SelectBill(constraints)

    def SelectBill(self, constraints):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            s = ""
            if constraints[0] == "p_name":
                cursor.execute(
                    "Select p_id FROM purchaser WHERE p_name =?;", (constraints[1],))
                rows = cursor.fetchall()
                p_id = rows[0][0]
                if len(constraints) == 2:
                    cursor.execute(
                        "Select pb_no, pb_year, pb_total_amt, pb_status FROM purchase_bill WHERE p_id=?;", (p_id,))
                elif len(constraints) == 3:
                    cursor.execute("Select pb_no, pb_year, pb_total_amt, pb_status FROM purchase_bill WHERE p_id=? AND pb_year=?;",
                                   (p_id, constraints[2],))
                else:
                    cursor.execute("Select pb_no, pb_year, pb_total_amt, pb_status FROM purchase_bill WHERE p_id=? AND pb_year=? AND pb_month=?;",
                                   (p_id, constraints[2], constraints[3]))
                rows = cursor.fetchall()
                for row in rows:
                    s += "\n"+str(row[0]).center(7)
                    s += str(row[1]).center(8)
                    s += constraints[1].center(47)
                    s += str(row[2]).center(13)
                    s += str(row[3]).center(17)
            else:
                if len(constraints) == 1:
                    cursor.execute("Select pb_no, pb_year, p_id, pb_total_amt, pb_status FROM purchase_bill WHERE pb_year=?;",
                                   (constraints[0],))
                else:
                    cursor.execute("Select pb_no, pb_year, p_id, pb_total_amt, pb_status FROM purchase_bill WHERE pb_year=? AND pb_month=?;",
                                   (constraints[0], constraints[1]))
                rows = cursor.fetchall()
                for row in rows:
                    s += "\n"+str(row[0]).center(7)
                    s += str(row[1]).center(8)
                    cursor.execute(
                        "Select p_name FROM purchaser WHERE p_id =?;", (row[2],))
                    p_name = cursor.fetchall()
                    s += p_name[0][0].center(47)
                    s += str(row[3]).center(13)
                    s += str(row[4]).center(17)
            self.displayText.configure(state='normal')
            self.displayText.delete('1.0', END)
            self.displayText.insert(INSERT, s)
            self.displayText.configure(state='disabled')
            sqliteConnection.commit()
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to Get Data form Database", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def generatePurchaseBill(self):
        year = int(self.byear.get())
        if not year:
            messagebox.showerror(
                title="Error", message="Enter a Year!!")
            return
        if not (2019 < year < 2030):
            messagebox.showerror(
                title="Error", message="Incorrect Year!!")
            return
        month = int(self.bmonth.get())
        if not month:
            messagebox.showerror(
                title="Error", message="Enter a Year!!")
            return
        if not (0 < month < 13):
            messagebox.showerror(
                title="Error", message="Incorrect Year!!")
            return
        Path("Bills/Purchases/"+str(year)
             ).mkdir(parents=True, exist_ok=True)
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()
            cursor.execute(
                "Select pb_day, pb_month, pb_year, pb_bill_no, pb_tax_amt, pb_gst_5, pb_igst_5, pb_gst_12, pb_igst_12, pb_gst_18, pb_igst_18, pb_gst_28, pb_igst_28, pb_total_amt, pb_status, p_id FROM purchase_bill WHERE pb_year=? AND pb_month=?;", (year, month,))
            rows = cursor.fetchall()
            workbook = xlsxwriter.Workbook(
                'Bills/Purchases/' + str(year) + '/Purchase Bill [' + months[month] + "-" + str(year)+'].xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.set_paper(9)
            worksheet.set_landscape()
            worksheet.center_horizontally()
            worksheet.set_margins(0.05, 0.05, 0.05, 0.05)
            worksheet.set_default_row(15)
            table_header = workbook.add_format({
                'border': 1,
                'bold': 'bold',
                'align': 'center',
                'valign': 'bottom',
                'font_name': 'Times New Roman',
                'font_size': 12})
            table_data = workbook.add_format({
                'border': 1,
                'valign': 'bottom',
                'align': 'center',
                'font_name': 'Times New Roman',
                'font_size': 11})

            # ------ Set column Width -----
            # Date
            worksheet.set_column(0, 0, 10)
            # Bill No.
            worksheet.set_column(1, 1, 10)
            # Purchaser
            worksheet.set_column(2, 2, 30)
            # GST
            worksheet.set_column(3, 3, 14)
            # Taxable Amount
            worksheet.set_column(4, 4, 16)
            # CGST-5
            worksheet.set_column(5, 5, 8)
            # SGST-5
            worksheet.set_column(6, 6, 8)
            # IGST-5
            worksheet.set_column(7, 7, 8)
            # CGST-12
            worksheet.set_column(8, 8, 8)
            # SGST-12
            worksheet.set_column(9, 9, 8)
            # IGST-12
            worksheet.set_column(10, 10, 8)
            # CGST-18
            worksheet.set_column(11, 11, 8)
            # SGST-18
            worksheet.set_column(12, 12, 8)
            # IGST-18
            worksheet.set_column(13, 13, 8)
            # CGST-28
            worksheet.set_column(14, 14, 8)
            # SGST-28
            worksheet.set_column(15, 15, 8)
            # IGST-28
            worksheet.set_column(16, 16, 8)
            # Total Amount
            worksheet.set_column(17, 17, 16)
            # Status Pending
            worksheet.set_column(18, 18, 8)

            # -------- Excelsheet --------

            # Table Columns
            worksheet.write(
                'A' + str(1), "Date", table_header)
            worksheet.write(
                'B' + str(1), "Bill No.", table_header)
            worksheet.write(
                'C' + str(1), "Purchaser", table_header)
            worksheet.write(
                'D' + str(1), "GST No.", table_header)
            worksheet.write(
                'E' + str(1), "Taxable Amount", table_header)
            worksheet.write(
                'F' + str(1), "cgst 5", table_header)
            worksheet.write(
                'G' + str(1), "sgst 5", table_header)
            worksheet.write(
                'H' + str(1), "igst 5", table_header)
            worksheet.write(
                'I' + str(1), "cgst 12", table_header)
            worksheet.write(
                'J' + str(1), "sgst 12", table_header)
            worksheet.write(
                'K' + str(1), "igst 12", table_header)
            worksheet.write(
                'L' + str(1), "cgst 18", table_header)
            worksheet.write(
                'M' + str(1), "sgst 18", table_header)
            worksheet.write(
                'N' + str(1), "igst 18", table_header)
            worksheet.write(
                'O' + str(1), "cgst 28", table_header)
            worksheet.write(
                'P' + str(1), "sgst 28", table_header)
            worksheet.write(
                'Q' + str(1), "igst 28", table_header)
            worksheet.write(
                'R' + str(1), "Total Amount", table_header)
            worksheet.write(
                'S' + str(1), "Status", table_header)

            for ent in range(0, len(rows)):
                cursor.execute(
                    "Select p_name, p_gst FROM purchaser WHERE p_id=?", (rows[ent][-1],))
                purchaser = cursor.fetchone()
                worksheet.write_string(
                    'A' + str(2 + ent), str(rows[ent][0]) + "/"+str(rows[ent][1]) + "/"+str(rows[ent][2]), table_data)
                worksheet.write_string(
                    'B' + str(2 + ent), rows[ent][3], table_data)
                worksheet.write_string(
                    'C' + str(2 + ent), purchaser[0], table_data)
                worksheet.write_string(
                    'D' + str(2 + ent), purchaser[1], table_data)
                worksheet.write_number(
                    'E' + str(2 + ent), rows[ent][4], table_data)
                worksheet.write_number(
                    'F' + str(2 + ent), rows[ent][5], table_data)
                worksheet.write_number(
                    'G' + str(2 + ent), rows[ent][5], table_data)
                worksheet.write_number(
                    'H' + str(2 + ent), rows[ent][6], table_data)
                worksheet.write_number(
                    'I' + str(2 + ent), rows[ent][7], table_data)
                worksheet.write_number(
                    'J' + str(2 + ent), rows[ent][7], table_data)
                worksheet.write_number(
                    'K' + str(2 + ent), rows[ent][8], table_data)
                worksheet.write_number(
                    'L' + str(2 + ent), rows[ent][9], table_data)
                worksheet.write_number(
                    'M' + str(2 + ent), rows[ent][9], table_data)
                worksheet.write_number(
                    'N' + str(2 + ent), rows[ent][10], table_data)
                worksheet.write_number(
                    'O' + str(2 + ent), rows[ent][11], table_data)
                worksheet.write_number(
                    'P' + str(2 + ent), rows[ent][11], table_data)
                worksheet.write_number(
                    'Q' + str(2 + ent), rows[ent][12], table_data)
                worksheet.write_number(
                    'R' + str(2 + ent), rows[ent][13], table_data)
                worksheet.write_string(
                    'S' + str(2 + ent), rows[ent][14], table_data)
            workbook.close()

            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Excel Bill created successfully!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to get Bill details from table", message=error)
        except xlsxwriter.exceptions.FileCreateError as error:
            messagebox.showerror(
                title="Failed to generate Bill", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def editStatus(self):
        b_year = self.byear.get()
        if not b_year:
            messagebox.showerror(
                title="Error", message="Billing Year Field cannot be empty!!")
            return
        b_no = self.bno.get()
        if not b_no:
            messagebox.showerror(
                title="Error", message="Billing No. Field cannot be empty!!")
            return
        status = self.estatus.get()
        if not status:
            messagebox.showerror(
                title="Error", message="Payment Status Field cannot be empty!!")
            return
        self.updateStatus((status, b_year, b_no,))

    def updateStatus(self, constraints):
        try:
            sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
            cursor = sqliteConnection.cursor()

            cursor.execute(
                "UPDATE purchase_bill SET pb_status=? WHERE pb_year=? AND pb_no=?;", constraints)
            messagebox.showinfo(
                title="Successfull", message="Payment Status updated Successfully!!")
            sqliteConnection.commit()
            cursor.close()
            self.SelectBill([constraints[1]])

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to Get Data form Database", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()

    def updatePurchaserList(self):
        getPurchaserList()
        self.pnameC['values'] = purchasers


# ------ Autocomplete Entry Class -------


class AutocompleteEntry(Entry):
    def __init__(self, autocompleteList, *args, **kwargs):

        # Listbox length
        if 'listboxLength' in kwargs:
            self.listboxLength = kwargs['listboxLength']
            del kwargs['listboxLength']
        else:
            self.listboxLength = 8

        # Custom matches function
        if 'matchesFunction' in kwargs:
            self.matchesFunction = kwargs['matchesFunction']
            del kwargs['matchesFunction']
        else:
            def matches(fieldValue, acListEntry):
                pattern = re.compile(
                    '.*' + re.escape(fieldValue) + '.*', re.IGNORECASE)
                return re.match(pattern, acListEntry)

            self.matchesFunction = matches

        Entry.__init__(self, *args, **kwargs)
        self.focus()

        self.autocompleteList = autocompleteList

        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()

        self.var.trace('w', self.changed)
        self.bind("<Return>", self.selection)
        self.bind("<Up>", self.moveUp)
        self.bind("<Down>", self.moveDown)

        self.listboxUp = False

    def changed(self, name, index, mode):
        if self.var.get() == '':
            if self.listboxUp:
                self.listbox.destroy()
                self.listboxUp = False
        else:
            words = self.comparison()
            if words:
                if not self.listboxUp:
                    self.listbox = Listbox(
                        width=40, height=self.listboxLength, font=("arial", 11))
                    self.listbox.bind("<Button-1>", self.selection)
                    self.listbox.bind("<Return>", self.selection)
                    self.listbox.place(relx=0.09, rely=0.35)
                    self.listboxUp = True

                self.listbox.delete(0, END)
                for w in words:
                    self.listbox.insert(END, w)
            else:
                if self.listboxUp:
                    self.listbox.destroy()
                    self.listboxUp = False

    def selection(self, event):
        if self.listboxUp:
            self.var.set(self.listbox.get(ACTIVE))
            self.listbox.destroy()
            self.listboxUp = False
            self.icursor(END)

    def moveUp(self, event):
        if self.listboxUp:
            if self.listbox.curselection() == ():
                index = '0'
            else:
                index = self.listbox.curselection()[0]

            if index != '0':
                self.listbox.selection_clear(first=index)
                index = str(int(index) - 1)

                self.listbox.see(index)  # Scroll!
                self.listbox.selection_set(first=index)
                self.listbox.activate(index)

    def moveDown(self, event):
        if self.listboxUp:
            if self.listbox.curselection() == ():
                index = '0'
            else:
                index = self.listbox.curselection()[0]

            if index != END:
                self.listbox.selection_clear(first=index)
                index = str(int(index) + 1)

                self.listbox.see(index)  # Scroll!
                self.listbox.selection_set(first=index)
                self.listbox.activate(index)

    def comparison(self):
        return [w for w in self.autocompleteList if self.matchesFunction(self.var.get(), w)]


class AutocompleteAddEntry(Entry):
    def __init__(self, autocompleteList, *args, **kwargs):

        # Listbox length
        if 'listboxLength' in kwargs:
            self.listboxLength = kwargs['listboxLength']
            del kwargs['listboxLength']
        else:
            self.listboxLength = 8

        # Custom matches function
        if 'matchesFunction' in kwargs:
            self.matchesFunction = kwargs['matchesFunction']
            del kwargs['matchesFunction']
        else:
            def matches(fieldValue, acListEntry):
                pattern = re.compile(
                    '.*' + re.escape(fieldValue) + '.*', re.IGNORECASE)
                return re.match(pattern, acListEntry)

            self.matchesFunction = matches

        Entry.__init__(self, *args, **kwargs)
        self.focus()

        self.autocompleteList = autocompleteList

        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()

        self.var.trace('w', self.changed)
        self.bind("<Return>", self.selection)
        self.bind("<Up>", self.moveUp)
        self.bind("<Down>", self.moveDown)

        self.listboxUp = False

    def changed(self, name, index, mode):
        if self.var.get() == '':
            if self.listboxUp:
                self.listbox.destroy()
                self.listboxUp = False
        else:
            words = self.comparison()
            if words:
                if not self.listboxUp:
                    self.listbox = Listbox(
                        width=33, height=self.listboxLength, font=("arial", 18))
                    self.listbox.bind("<Button-1>", self.selection)
                    self.listbox.bind("<Return>", self.selection)
                    self.listbox.place(relx=0.11, rely=0.31)
                    self.listboxUp = True

                self.listbox.delete(0, END)
                for w in words:
                    self.listbox.insert(END, w)
            else:
                if self.listboxUp:
                    self.listbox.destroy()
                    self.listboxUp = False

    def selection(self, event):
        if self.listboxUp:
            self.var.set(self.listbox.get(ACTIVE))
            self.listbox.destroy()
            self.listboxUp = False
            self.icursor(END)

    def moveUp(self, event):
        if self.listboxUp:
            if self.listbox.curselection() == ():
                index = '0'
            else:
                index = self.listbox.curselection()[0]

            if index != '0':
                self.listbox.selection_clear(first=index)
                index = str(int(index) - 1)

                self.listbox.see(index)  # Scroll!
                self.listbox.selection_set(first=index)
                self.listbox.activate(index)

    def moveDown(self, event):
        if self.listboxUp:
            if self.listbox.curselection() == ():
                index = '0'
            else:
                index = self.listbox.curselection()[0]

            if index != END:
                self.listbox.selection_clear(first=index)
                index = str(int(index) + 1)

                self.listbox.see(index)  # Scroll!
                self.listbox.selection_set(first=index)
                self.listbox.activate(index)

    def comparison(self):
        return [w for w in self.autocompleteList if self.matchesFunction(self.var.get(), w)]

def matches(fieldValue, acListEntry):
    pattern = re.compile(re.escape(fieldValue) + '.*', re.IGNORECASE)
    return re.match(pattern, acListEntry)

# ---------- Global Update List Function --------


def getCientList():
    sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
    cursor = sqliteConnection.cursor()
    cursor.execute("""Select c_name FROM client;""")
    rows = cursor.fetchall()
    clients.clear()
    for i in range(len(rows)):
        clients.append(rows[i][0])
    clients.sort()
    sqliteConnection.commit()
    cursor.close()
    if (sqliteConnection):
        sqliteConnection.close()


def getPurchaserList():
    sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
    cursor = sqliteConnection.cursor()
    cursor.execute("""Select p_name FROM purchaser;""")
    rows = cursor.fetchall()
    purchasers.clear()
    for i in range(len(rows)):
        purchasers.append(rows[i][0])
    purchasers.sort()
    sqliteConnection.commit()
    cursor.close()
    if (sqliteConnection):
        sqliteConnection.close()


def addProduct(name):
    sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
    cursor = sqliteConnection.cursor()
    cursor.execute("INSERT INTO product VALUES (?)", (name,))
    sqliteConnection.commit()
    cursor.close()
    if (sqliteConnection):
        sqliteConnection.close()
    getProductList()


def getProductList():
    sqliteConnection = sqlite3.connect('Bills/Database/Billing.db')
    cursor = sqliteConnection.cursor()
    cursor.execute("""Select pr_name FROM product;""")
    rows = cursor.fetchall()
    products.clear()
    for i in range(len(rows)):
        products.append(rows[i][0])
    products.sort()
    sqliteConnection.commit()
    cursor.close()
    if (sqliteConnection):
        sqliteConnection.close()


# ---------- Number to Word -----------

def numToWords(n, s):
    str = ""
    if (n > 19):
        str += ten[n // 10] + one[n % 10]
    else:
        str += one[n]
    if (n):
        str += s
    return str


def convertToWords(n):
    out = ""
    out += numToWords((n // 10000000), "Crore ")
    out += numToWords(((n // 100000) % 100), "Lakh ")
    out += numToWords(((n // 1000) % 100), "Thousand ")
    out += numToWords(((n // 100) % 10), "Hundred ")
    if (n > 100 and n % 100):
        out += "and "
    out += numToWords((n % 100), "")
    return out


if __name__ == "__main__":
    getProductList()
    app = App()
    app.mainloop()

