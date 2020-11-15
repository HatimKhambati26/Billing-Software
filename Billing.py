from tkinter import *
from tkinter import messagebox, ttk, scrolledtext
import sqlite3
from string import ascii_uppercase as alphabet
charcoal = '#F1F1EE'
rust = '#F62A00'
navy = '#00293C'
teal = '#1E656D'
clients = []
purchasers = []
products = ["Some Product-1", "Some Product-2", "Some Product-3",
            "Some Product-4", "Some Product-5", "Some Product-6", "Some Product-7", "Some Product-2", "Some Product-3",
            "Some Product-4", "Some Product-5", "Some Product-6", "Some Product-7"]
years = ["2020-2021", "2021-2022", "2022-2023",  "2023-2024",
         "2024-2025", "2025-2026", "2026-2027", "2027-2028", "2028-2029", "2029-2030"]
months = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
days = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16",
        "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"]


class App(Tk):

    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        container = Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        self.tk_setPalette(background=charcoal, foreground=navy,
                           activeBackground=rust, activeForeground=charcoal)
        self.geometry("1920x1080+0+0")
        self.title("Billing Software")
        self.frames = {}
        for F in (Home, AddClient, CreateBill, AddBillDetails, UpdateBillStatus, AddPurchaseBill, UpdatePurchaseStatus):
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
            "times new roman", 42, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=460, pady=60)
        saleF.place(x=1, y=130, relwidth=1, height=420)

        # Create Sales Bill
        Button(saleF, text="Create New Bill", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=32, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(CreateBill)).grid(row=0, column=0, padx=20, pady=15)

        # Edit Sales Bill Button
        Button(saleF, text="Edit Bill", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=32, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddBillDetails)).grid(row=0, column=1, padx=20, pady=15)

        # Add Client Button
        Button(saleF, text="Add New Client", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=32, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddClient)).grid(row=1, column=0, padx=20, pady=15)

        # Check & Edit Bill Status Button
        Button(saleF, text="Search & Edit Bill Status", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=32, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(UpdateBillStatus)).grid(row=1, column=1, padx=20, pady=15)

        # --------- Purchase Options ---------
        purchaseF = LabelFrame(self, text="Purchase", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=460, pady=100)
        purchaseF.place(x=1, y=600, relwidth=1, height=425)

        # Create Purchase Bill
        Button(purchaseF, text="Create New Bill", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=32, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddPurchaseBill)).grid(row=0, column=0, padx=20, pady=15)

        # Search Bill Button
        Button(purchaseF, text="Search & Edit Bill Status", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=32, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(UpdatePurchaseStatus)).grid(row=0, column=2, padx=20, pady=15)


# ---------- SALES FRAMES --------


class AddClient(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 42, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=220, pady=60)
        self.saleF.place(x=1, y=130, relwidth=1, height=895)

        Label(self.saleF, text="Client's Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold"), width=85, pady=10).grid(row=0, column=0, columnspan=2, pady=15)

        # Client's Name
        Label(self.saleF, text="Client's Name:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=1, column=0, pady=15, sticky=E)
        self.cname_txt = Entry(self.saleF, width=40, bd=3, relief=GROOVE, font=(
            "arial", 22, "bold"))
        self.cname_txt.grid(row=1, column=1, padx=50, pady=15, sticky=W)

        # Client's GST No
        Label(self.saleF, text="Client's GST No:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=2, column=0, pady=15, sticky=E)
        self.cgst_txt = Entry(self.saleF, width=40, bd=3, relief=GROOVE, font=(
            "arial", 22, "bold"))
        self.cgst_txt.grid(row=2, column=1, padx=50, pady=15, sticky=W)

        # Client's Address
        Label(self.saleF, text="Client's Address:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=3, column=0, pady=15, sticky=E)
        self.caddress_txt = Entry(self.saleF, width=40, bd=3, relief=GROOVE, font=(
            "arial", 22, "bold"))
        self.caddress_txt.grid(row=3, column=1, padx=50, pady=15, sticky=W)

        # Add Client Details Button
        add_client_btn = Button(self.saleF, text="Add New Client", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=50, font=(
            "arial", 18, "bold"), command=self.addClient).place(x=400, y=400)

        # CLear Details Button
        clear_btn = Button(self.saleF, text="Clear", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=self.clearText).place(x=400, y=550)

        # Home Button
        add_client_btn = Button(self.saleF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(x=790, y=550)

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
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")

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
                # print("The SQLite connection is closed")

    def clearText(self):
        self.cname_txt.delete(0, END)
        self.cgst_txt.delete(0, END)
        self.caddress_txt.delete(0, END)


class CreateBill(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 42, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=180, pady=60)
        self.saleF.place(x=1, y=130, relwidth=1, height=895)

        Label(self.saleF, text="Bill Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold"), width=85, pady=10).grid(row=0, column=0, columnspan=6, pady=15)

        # Billing Client's Name
        Label(self.saleF, text="Client's Name:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=1, column=0, columnspan=2, pady=15, sticky=E)
        self.cname = StringVar()
        self.cname_txt = ttk.Combobox(self.saleF, width=40, font=(
            "arial", 22, "bold"), textvariable=self.cname, postcommand=self.updateClientList)
        self.cname_txt.grid(row=1, column=2, columnspan=4,
                            padx=50, pady=25, sticky=W)

        # Billing Year
        Label(self.saleF, text="Financial Year:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=2, column=0, pady=15, sticky=E)
        self.byear = StringVar()
        self.byear_txt = OptionMenu(self.saleF, self.byear, *years)
        self.byear_txt.config(width=15, font=(
            "arial", 22, "bold"), bd=3, relief=GROOVE)
        self.byear_txt.grid(row=2, column=1, padx=50, pady=15, sticky=W)

        # Billing Date
        Label(self.saleF, text="Date:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=2, column=2, pady=15, sticky=E)
        self.bdate = Entry(self.saleF, width=15, font=(
            "arial", 22, "bold"), bd=3, relief=GROOVE)
        self.bdate.grid(row=2, column=3, padx=50, pady=15, sticky=W)

        # P.O. Number
        Label(self.saleF, text="P.O. Number.:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=2, column=4, pady=15, sticky=E)
        self.bpo = Entry(self.saleF, width=15, font=(
            "arial", 22, "bold"), bd=3, relief=GROOVE)
        self.bpo.grid(row=2, column=5, padx=50, pady=65, sticky=W)

        # Add Bill Button
        Button(self.saleF, text="Add Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=50, font=(
            "arial", 18, "bold"), command=self.createBill).place(x=410, y=400)

        # Add Bill Details Button
        clear_btn = Button(self.saleF, text="Add Bill Details", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(AddBillDetails)).place(x=410, y=550)

        # Home Button
        add_client_btn = Button(self.saleF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=20, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(x=800, y=550)

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
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute(
                """SELECT MAX(b_no) FROM bill WHERE b_year=?;""", (year,))
            val = cursor.fetchall()
            if val[0][0] == None:
                bill_no = 1
            else:
                bill_no = val[0][0]+1
            cursor.execute(
                """SELECT c_id FROM client WHERE c_name=?;""", (name,))
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
                # print("The SQLite connection is closed")

    def updateClientList(self):
        getCientList()
        self.cname_txt['values'] = clients


class AddBillDetails(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 42, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=20, pady=10)
        self.saleF.place(x=1, y=80, relwidth=1, height=945)

        Label(self.saleF, text="Bill Product Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold"), width=85, pady=10).place(x=1, y=1, relwidth=1)

        # ---------- Add Details Frame ------------
        self.addDetailsF = LabelFrame(self.saleF, text="Add Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=10, pady=20)
        self.addDetailsF.place(x=1, y=100, width=615, height=400)

        # Product Name
        Label(self.addDetailsF, text="Product:", font=(
            "times new roman", 14, "bold")).grid(row=0, column=0, pady=10, sticky=E)
        self.pname = StringVar()
        self.pnameC = ttk.Combobox(self.addDetailsF, width=15, font=(
            "arial", 14, "bold"), textvariable=self.pname, postcommand=self.updateProductList)
        self.pnameC.grid(row=0, column=1, padx=20, pady=10)

        # Quantity
        Label(self.addDetailsF, text="Quantity:", font=(
            "times new roman", 14, "bold")).grid(row=0, column=2, pady=10, sticky=E)
        self.pquan = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.pquan.grid(row=0, column=3, padx=10, pady=10)

        # Rate
        Label(self.addDetailsF, text="Rate:", font=(
            "times new roman", 14, "bold")).grid(row=1, column=0, pady=10, sticky=E)
        self.rate = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.rate.grid(row=1, column=1, padx=10, pady=10)

        # Taxable Amount
        Label(self.addDetailsF, text="Amount:", font=(
            "times new roman", 14, "bold")).grid(row=1, column=2, pady=10, sticky=E)
        self.amt = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.amt.grid(row=1, column=3, padx=10, pady=10)

        # GST
        self.gst = DoubleVar()
        # CGST
        Label(self.addDetailsF, text="CGST:", font=(
            "times new roman", 14, "bold")).grid(row=2, column=0, pady=10, sticky=E)
        Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"), textvariable=self.gst).grid(row=2, column=1, padx=10, pady=10)

        # SGST
        Label(self.addDetailsF, text="SGST:", font=(
            "times new roman", 14, "bold")).grid(row=2, column=2, pady=10, sticky=E)
        Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"), textvariable=self.gst).grid(row=2, column=3, padx=10, pady=10)

        # IGST
        Label(self.addDetailsF, text="IGST:", font=(
            "times new roman", 14, "bold")).grid(row=3, column=0, pady=10, sticky=E)
        self.igst = DoubleVar()
        Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"), textvariable=self.igst).grid(row=3, column=1, padx=10, pady=10)

        # Challan No.
        Label(self.addDetailsF, text="Challan No:", font=(
            "times new roman", 14, "bold")).grid(row=3, column=2, pady=10, sticky=E)
        self.challan = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.challan.grid(row=3, column=3, padx=10, pady=10)

        # HSN Code
        Label(self.addDetailsF, text="HSN Code:", font=(
            "times new roman", 14, "bold")).grid(row=4, column=0, pady=10, sticky=E)
        self.phsn = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.phsn.grid(row=4, column=1, padx=10, pady=10)

        # AMC No.
        Label(self.addDetailsF, text="AMC No:", font=(
            "times new roman", 14, "bold")).grid(row=4, column=2, pady=10, sticky=E)
        self.amc = Entry(self.addDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.amc.grid(row=4, column=3, padx=10, pady=10)

        # Insert Bill Details Button
        Button(self.addDetailsF, text="Add Bill Detail", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=15, font=(
            "arial", 16, "bold"), command=self.addBillDetails).grid(row=5, column=0, columnspan=2, padx=10, pady=10)

        # Clear Details Button
        Button(self.addDetailsF, text="Clear", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=15, font=(
            "arial", 16, "bold"), command=self.clearTextAdd).grid(row=5, column=2, columnspan=2, padx=10, pady=10)

        # --------- Edit Details Frame ------------
        self.editDetailsF = LabelFrame(self.saleF, text="Edit Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=10, pady=20)
        self.editDetailsF.place(x=645, y=100, width=615, height=400)

        # Bill Serial No.
        Label(self.editDetailsF, text="", font=(
            "times new roman", 14, "bold")).grid(row=0, column=0, columnspan=4, pady=10, sticky=E)

        Label(self.editDetailsF, text="Serial No.:", font=(
            "times new roman", 14, "bold")).place(x=1, y=10)
        self.esrno = StringVar()
        self.esrnoC = Entry(self.editDetailsF, width=4, font=(
            "arial", 14, "bold"), textvariable=self.esrno, bd=3, relief=GROOVE)
        self.esrnoC.place(x=100, y=10)

        # HSN Code
        Label(self.editDetailsF, text="HSN Code:", font=(
            "times new roman", 14, "bold")).place(x=180, y=10)
        self.ephsn = Entry(self.editDetailsF, width=6, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.ephsn.place(x=280, y=10)

        # AMC No.
        Label(self.editDetailsF, text="AMC No:", font=(
            "times new roman", 14, "bold")).place(x=380, y=10)
        self.eamc = Entry(self.editDetailsF, width=7, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.eamc.place(x=470, y=10)

        # Product Name
        Label(self.editDetailsF, text="Product:", font=(
            "times new roman", 14, "bold")).grid(row=1, column=0, pady=10, sticky=E)
        self.epname = StringVar()
        self.epnameC = ttk.Combobox(self.editDetailsF, width=15, font=(
            "arial", 14, "bold"), textvariable=self.epname, postcommand=self.updateProductList)
        self.epnameC.grid(row=1, column=1, padx=20, pady=10)

        # Quantity
        Label(self.editDetailsF, text="Quantity:", font=(
            "times new roman", 14, "bold")).grid(row=1, column=2, pady=10, sticky=E)
        self.epquan = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.epquan.grid(row=1, column=3, padx=10, pady=10)

        # Rate
        Label(self.editDetailsF, text="Rate:", font=(
            "times new roman", 14, "bold")).grid(row=2, column=0, pady=10, sticky=E)
        self.erate = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.erate.grid(row=2, column=1, padx=10, pady=10)

        # Taxable Amount
        Label(self.editDetailsF, text="Amount:", font=(
            "times new roman", 14, "bold")).grid(row=2, column=2, pady=10, sticky=E)
        self.eamt = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.eamt.grid(row=2, column=3, padx=10, pady=10)

        # CGST
        Label(self.editDetailsF, text="CGST:", font=(
            "times new roman", 14, "bold")).grid(row=3, column=0, pady=10, sticky=E)
        self.ecgst = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.ecgst.grid(row=3, column=1, padx=10, pady=10)

        # SGST
        Label(self.editDetailsF, text="SGST:", font=(
            "times new roman", 14, "bold")).grid(row=3, column=2, pady=10, sticky=E)
        self.esgst = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.esgst.grid(row=3, column=3, padx=10, pady=10)

        # IGST
        Label(self.editDetailsF, text="IGST:", font=(
            "times new roman", 14, "bold")).grid(row=4, column=0, pady=10, sticky=E)
        self.eigst = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.eigst.grid(row=4, column=1, padx=10, pady=10)

        # Challan No.
        Label(self.editDetailsF, text="Challan No:", font=(
            "times new roman", 14, "bold")).grid(row=4, column=2, pady=10, sticky=E)
        self.echallan = Entry(self.editDetailsF, width=15, bd=3, relief=GROOVE, font=(
            "arial", 14, "bold"))
        self.echallan.grid(row=4, column=3, padx=10, pady=10)

        # Edit Bill Details Button
        Button(self.editDetailsF, text="Edit Bill Detail", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=15, font=(
            "arial", 16, "bold"), command=self.editBillDetails).grid(row=5, column=0, columnspan=2, padx=10, pady=10)

        # Clear Edit Details Button
        Button(self.editDetailsF, text="Clear", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=15, font=(
            "arial", 16, "bold"), command=self.clearTextEdit).grid(row=5, column=2, columnspan=2, padx=10, pady=10)

        # --------- Select Bill Frame ------------
        self.selectBillF = LabelFrame(self.saleF, text="Select Bill", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=20, pady=10)
        self.selectBillF.place(x=1290, y=100, width=580, height=180)

        # Billing Year
        Label(self.selectBillF, text="Billing Year:", font=(
            "times new roman", 16, "bold")).grid(row=0, column=0, pady=10, sticky=E)
        self.byear = StringVar()
        self.byearC = ttk.Combobox(self.selectBillF, width=10, font=(
            "arial", 16, "bold"), textvariable=self.byear, postcommand=self.updateYearList)
        self.byearC.grid(row=0, column=1, padx=10, pady=10)

        # Bill No.
        Label(self.selectBillF, text="Billing No:", font=(
            "times new roman", 16, "bold")).grid(row=0, column=2, pady=10, sticky=E)
        self.bno = StringVar()
        self.bnoC = Entry(self.selectBillF, width=10, bd=3, relief=GROOVE, font=(
            "arial", 16, "bold"), textvariable=self.bno)
        self.bnoC.grid(row=0, column=3, padx=10, pady=10)

        # Current Bill Details Button
        Button(self.selectBillF, text="View Bill Details", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=18, font=(
            "arial", 16, "bold"), command=self.viewBillDetails).place(x=1, y=60)

        # Delete Last Entry Bill Details Button
        Button(self.selectBillF, text="Delete Last Entry", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=18, font=(
            "arial", 16, "bold"), command=self.removeLastDetail).place(x=270, y=60)

        # --------- Edit Client Frame ------------
        self.editClientF = LabelFrame(self.saleF, text="Edit Client", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=30, pady=10)
        self.editClientF.place(x=1290, y=290, width=580, height=210)

        # Billing Client's Name
        Label(self.editClientF, text="Client's Name:", font=(
            "times new roman", 16, "bold")).grid(row=0, column=0, pady=20, sticky=E)
        self.ecname = StringVar()
        self.cname_txt = ttk.Combobox(self.editClientF, width=25, font=(
            "arial", 16, "bold"), textvariable=self.ecname, postcommand=self.updateClientList)
        self.cname_txt.grid(row=0, column=1,
                            padx=20, pady=10, sticky=W)

        # Edit Client Button
        Button(self.editClientF, text="Edit Client", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=15, pady=10, font=(
            "arial", 16, "bold"), command=self.editClient).place(x=150, y=80)

        # --------- View Bill Detials Frame ------------
        self.viewBillF = LabelFrame(self.saleF, text="View Bill Detials", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=10, pady=10)
        self.viewBillF.place(x=1, y=512, width=1355, height=360)
        self.displayBillF = Frame(self.viewBillF)
        self.displayBillF.place(x=10, y=20, width=1305, height=280)
        self.displayText = scrolledtext.ScrolledText(
            self.displayBillF, font=("Courier",
                                     12, "bold"), padx=10, pady=10)
        self.displayText.insert(INSERT,
                                "\n\n\n\n\n\n\n\n\t\t\t\t\tSelect Billing Year and Billing No. to see the entries!! ")
        self.displayText.configure(state='disabled')
        self.displayText.pack(side="left", fill="both", expand=True)
        Label(self.viewBillF, text="Sr. No.\t\t            Product\t\t              Rate      Quantity CGST   SGST    IGST       AMOUNT\t     HSN         Ch. No.      AMC No.", font=(
            "times new roman", 15, "bold")).place(x=30, y=0)

        # --------- Bill Information Frame ------------
        self.viewInfoF = LabelFrame(self.saleF, text="Bill Information", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=10, pady=10)
        self.viewInfoF.place(x=1380, y=512, width=490, height=230)
        self.displayInfoF = Frame(self.viewInfoF)
        self.displayInfoF.place(x=10, y=10, width=440, height=150)
        self.displayInfo = scrolledtext.ScrolledText(
            self.displayInfoF, font=("arial",
                                     16, "bold"), padx=10, pady=10)
        self.displayInfo.insert(INSERT,
                                "\n       Select Billing Year and Billing No. \n              to see the Information!! ")
        self.displayInfo.configure(state='disabled')
        self.displayInfo.pack(side="left", fill="both", expand=True)

        # Update Date Button
        Button(self.saleF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=35, pady=30, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(x=1380, y=770)

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
        if name not in products:
            messagebox.showerror(
                title="Error", message="Select Product's Name from Dropdown List!!")
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
        rate = self.rate.get()
        if not rate:
            messagebox.showerror(
                title="Error", message="Rate Field cannot be empty!!")
            return
        if not str.isdigit(rate):
            messagebox.showerror(
                title="Error", message="Enter a Valid Rate!!")
            return
        amt = self.amt.get()
        if not amt:
            messagebox.showerror(
                title="Error", message="Taxable Amount Field cannot be empty!!")
            return
        if not str.isdigit(amt):
            messagebox.showerror(
                title="Error", message="Enter a Valid Taxable Amount!!")
            return
        gst = self.gst.get()
        igst = self.igst.get()
        if not gst and not igst:
            messagebox.showerror(
                title="Error", message="Enter CGST/SGST or IGST!!")
            return
        hsn = self.phsn.get()
        amc = self.amc.get()
        challan = self.challan.get()

        data_list = [b_year, b_no, name, rate, quantity,
                     gst, gst, igst, amt, hsn, challan, amc]
        self.insertBillDetails(data_list)

    def insertBillDetails(self, data_list):
        try:
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute(
                "SELECT b_no FROM bill WHERE b_year=? AND b_no=?;", (data_list[0], data_list[1],))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            cursor.execute(
                """SELECT MAX(bd_id) FROM bill_detail WHERE b_year=? AND b_no=?;""", (data_list[0], data_list[1],))
            val = cursor.fetchall()
            if val[0][0] == None:
                id = 1
            else:
                id = val[0][0]+1
            data_list.insert(2, id)
            sqlite_insert_with_param = """INSERT INTO bill_detail
                            (b_year, b_no,bd_id, bd_product, bd_rate, bd_quantity,
                             bd_cgst, bd_sgst, bd_igst, bd_amount, bd_hsn, bd_ch_no, bd_amc_no)
                            VALUES (?, ?, ?,?, ?, ?,?, ?,?,?,?, ?, ?);"""
            cursor.execute(sqlite_insert_with_param, tuple(data_list))
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Details added successfully!!")
            cursor.close()
            self.viewBillDetails()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into bill_detail table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                # print("The SQLite connection is closed")

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
        rate = self.erate.get()
        if rate:
            if not str.isdigit(rate):
                messagebox.showerror(
                    title="Error", message="Enter a Valid Taxable Amount!!")
                return
            constraints.append("bd_rate")
            constraints.append(rate)
        amt = self.eamt.get()
        if amt:
            if not str.isdigit(amt):
                messagebox.showerror(
                    title="Error", message="Enter a Valid Taxable Amount!!")
                return
            constraints.append("bd_amount")
            constraints.append(amt)
        quantity = self.epquan.get()
        if quantity:
            if not str.isdigit(quantity):
                messagebox.showerror(
                    title="Error", message="Enter a Valid Quantity!!")
                return
            constraints.append("bd_quantity")
            constraints.append(quantity)
        cgst = self.ecgst.get()
        if cgst:
            constraints.append("bd_cgst")
            constraints.append(cgst)
        sgst = self.esgst.get()
        if sgst:
            constraints.append("bd_sgst")
            constraints.append(sgst)
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
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute(
                "SELECT b_no FROM bill WHERE b_year=? AND b_no=?;", (b_year, b_no,))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            data_tuple = (srno, b_year, b_no,)
            cursor.execute(
                "SELECT bd_id FROM bill_detail WHERE bd_id=? AND b_year=? AND b_no=?;", data_tuple)
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

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into Bill Details table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                # print("The SQLite connection is closed")

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
        self.selectBillDetails(b_year, b_no)
        self.selectBillInfo(b_year, b_no)

    def selectBillDetails(self, b_year, b_no):
        try:
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")

            sqlite_insert_with_param = """SELECT bd_id, bd_product, bd_rate, bd_quantity, bd_cgst, bd_sgst, bd_igst, bd_amount, bd_hsn, bd_ch_no, bd_amc_no FROM bill_detail WHERE b_year=? AND b_no=?;"""

            data_tuple = (b_year, b_no)
            cursor.execute(sqlite_insert_with_param, data_tuple)
            rows = cursor.fetchall()
            if not rows:
                cursor.execute(
                    "SELECT b_no FROM bill WHERE b_year=? AND b_no=?", data_tuple)
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
                s += str(row[4]).center(7)  # CGST
                s += str(row[5]).center(7)  # SGST
                s += str(row[6]).center(7)  # IGST
                s += str(row[7]).center(14)  # Amount
                s += str(row[8]).center(10)  # HSN
                s += str(row[9]).center(10)  # Challan
                s += str(row[10]).center(10)  # AMC
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
                # print("The SQLite connection is closed")

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
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute(
                "SELECT b_no FROM bill WHERE b_year=? AND b_no=?;", (b_year, b_no,))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showerror(
                    title="Error", message="Bill No. does not EXIST!!")
                return
            cursor.execute(
                "SELECT MAX(bd_id) FROM bill_detail WHERE b_year=? AND b_no=?;", (b_year, b_no,))
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
                # print("The SQLite connection is closed")

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
        self.updateClient(b_year, b_no, cname)

    def updateClient(self, b_year, b_no, cname):
        try:
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute("SELECT c_id FROM client WHERE c_name=?;", (cname,))
            c_id = cursor.fetchall()
            sqlite_insert_with_param = "UPDATE bill SET c_id = ? WHERE b_year = ? AND b_no = ?;"
            cursor.execute(
                sqlite_insert_with_param, (c_id[0][0], b_year, b_no,))
            sqliteConnection.commit()
            getCientList()
            cursor.close()
            self.selectBillInfo(b_year, b_no)

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into Bill table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                # print("The SQLite connection is closed")

    def selectBillInfo(self, b_year, b_no):
        try:
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")

            sqlite_insert_with_param = """SELECT c_id FROM bill WHERE b_year=? AND b_no=?;"""
            data_tuple = (b_year, b_no)
            cursor.execute(sqlite_insert_with_param, data_tuple)
            rows = cursor.fetchall()
            if not rows:
                return
            c_id = rows[0][0]
            sqlite_insert_with_param = """SELECT c_name FROM client WHERE c_id=?;"""
            cursor.execute(sqlite_insert_with_param, (c_id,))
            rows = cursor.fetchall()
            s = "Client's Name: "+rows[0][0]+"\n\n"
            sqlite_insert_with_param = """SELECT b_date, b_status, b_po FROM bill WHERE b_year=? AND b_no=?;"""
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
                # print("The SQLite connection is closed")

    def clearTextAdd(self):
        self.pnameC.delete(0, END)
        self.phsn.delete(0, END)
        self.pquan.delete(0, END)
        self.cgst.delete(0, END)
        self.sgst.delete(0, END)
        self.igst.delete(0, END)
        self.challan.delete(0, END)
        self.amc.delete(0, END)

    def clearTextEdit(self):
        self.esrnoC.delete(0, END)
        self.epnameC.delete(0, END)
        self.ephsn.delete(0, END)
        self.epquan.delete(0, END)
        self.ecgst.delete(0, END)
        self.esgst.delete(0, END)
        self.eigst.delete(0, END)
        self.echallan.delete(0, END)
        self.eamc.delete(0, END)

    def updateProductList(self):
        # getCientList()
        self.pnameC['values'] = products
        self.epnameC['values'] = products

    def updateYearList(self):
        # getCientList()
        self.byearC['values'] = years

    def updateClientList(self):
        getCientList()
        self.cname_txt['values'] = clients


class UpdateBillStatus(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 42, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.saleF = LabelFrame(self, text="Sales", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=220, pady=10)
        self.saleF.place(x=1, y=130, relwidth=1, height=895)

        Label(self.saleF, text="Check & Edit Bill Payment", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold"), width=85, pady=10).place(x=20, y=20, relwidth=1, height=60)

        # --------- Check Bill Status Frame ------------
        self.selectBillF = LabelFrame(self.saleF, text="  Search Bill Payment  ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"), padx=20, pady=20)
        self.selectBillF.place(x=30, y=100, width=700, height=320)

        # Billing Year
        Label(self.selectBillF, text="Billing Year:", font=(
            "times new roman", 18, "bold")).grid(row=0, column=0, padx=20, pady=10, sticky=E)
        self.byear = StringVar()
        self.byearC = ttk.Combobox(self.selectBillF, width=30, font=(
            "arial", 18, "bold"), textvariable=self.byear, postcommand=self.updateYearList)
        self.byearC.grid(row=0, column=1, padx=20, pady=10)

        # Client's Name
        Label(self.selectBillF, text="Client's Name:", font=(
            "times new roman", 18, "bold")).grid(row=1, column=0, padx=20, pady=20, sticky=E)
        self.cname = StringVar()
        self.cnameC = ttk.Combobox(self.selectBillF, width=30, font=(
            "arial", 18, "bold"), textvariable=self.cname, postcommand=self.updateClientList)
        self.cnameC.grid(row=1, column=1, padx=20, pady=10)

        # Search Bills Button
        Button(self.selectBillF, text="Search Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=30, font=(
            "arial", 18, "bold"), command=self.searchBill).place(x=120, y=150)

        # --------- Edit Bill Status Frame ------------
        self.editStatusF = LabelFrame(self.saleF, text="  Update Bill Payment  ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"), padx=20, pady=20)
        self.editStatusF.place(x=780, y=100, width=700, height=320)

        # Bill No.
        Label(self.editStatusF, text="Billing No:", font=(
            "times new roman", 18, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky=E)
        self.bno = StringVar()
        self.bnoC = Entry(self.editStatusF, width=30, font=(
            "arial", 18, "bold"), textvariable=self.bno, bd=3, relief=GROOVE)
        self.bnoC.grid(row=0, column=1, padx=10, pady=10)

        # Payment Status
        Label(self.editStatusF, text="Payment Status:", font=(
            "times new roman", 18, "bold")).grid(row=1, column=0, padx=15, pady=20, sticky=E)
        self.estatus = Entry(self.editStatusF, width=30, font=(
            "arial", 18, "bold"), bd=3, relief=GROOVE)
        self.estatus.grid(row=1, column=1, padx=10, pady=10)

        # Current Bill Details Button
        Button(self.editStatusF, text="Update Status", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=30, font=(
            "arial", 18, "bold"), command=self.editStatus).place(x=120, y=150)

        # --------- View Bill Detials Frame ------------
        self.viewBillF = LabelFrame(self.saleF, text="View Bill Detials", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=10, pady=10)
        self.viewBillF.place(x=30, y=440, width=990, height=360)
        self.displayBillF = Frame(self.viewBillF)
        self.displayBillF.place(x=10, y=20, width=940, height=280)
        self.displayText = scrolledtext.ScrolledText(
            self.displayBillF, font=("Courier",
                                     12, "bold"), padx=10, pady=10)
        self.displayText.insert(INSERT,
                                "\n\n\n\n\n\n\t\t\tSelect Billing Year or Client's Name to see the entries!! ")
        self.displayText.configure(state='disabled')
        self.displayText.pack(side="left", fill="both", expand=True)
        Label(self.viewBillF, text="Bill. No.\t       Year\t\t\t\t     Client's Name\t\t\t              Payment", font=(
            "times new roman", 15, "bold")).place(x=30, y=0)

        # Home Button
        Button(self.saleF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=25, pady=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(x=1080, y=600)

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
            # constraints.append("b_year")
            constraints.append(b_year)
        if not len(constraints):
            messagebox.showerror(
                title="Error", message="Select atleast 1 critera to search!!")
            return
        self.selectBill(constraints)

    def selectBill(self, constraints):
        try:
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            s = ""
            if constraints[0] == "c_name":
                cursor.execute(
                    "SELECT c_id FROM client WHERE c_name =?;", (constraints[1],))
                rows = cursor.fetchall()
                c_id = rows[0][0]
                if len(constraints) == 2:
                    cursor.execute(
                        "SELECT b_no, b_year, b_status FROM bill WHERE c_id=?;", (c_id,))
                    rows = cursor.fetchall()
                else:
                    cursor.execute("SELECT b_no, b_year, b_status FROM bill WHERE c_id=? AND b_year=?;",
                                   (c_id, constraints[2],))
                    rows = cursor.fetchall()
                for row in rows:
                    s += "\n"+str(row[0]).center(7)
                    s += str(row[1]).center(17)
                    s += constraints[1].center(51)
                    s += str(row[2]).center(17)
            else:
                cursor.execute("SELECT b_no, b_year,c_id, b_status FROM bill WHERE b_year=?;",
                               (constraints[0],))
                rows = cursor.fetchall()
                for row in rows:
                    s += "\n"+str(row[0]).center(7)
                    s += str(row[1]).center(17)
                    cursor.execute(
                        "SELECT c_name FROM client WHERE c_id =?;", (row[2],))
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
                # print("The SQLite connection is closed")

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
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute("SELECT b_no FROM bill WHERE b_year=? AND b_no=?", (
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
            self.selectBill([constraints[1]])

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to Get Data form Database", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                # print("The SQLite connection is closed")

    def updateYearList(self):
        # getCientList()
        self.byearC['values'] = years

    def updateClientList(self):
        getCientList()
        self.cnameC['values'] = clients


# ---------- PURCHASES FRAMES --------


class AddPurchaseBill(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        title = Label(self, text="F. K. PATANWALA & Co.", bd=8, relief=GROOVE, font=(
            "times new roman", 42, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Purchases Options ---------
        self.purchaseF = LabelFrame(self, text="Purchases", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=100, pady=20)
        self.purchaseF.place(x=1, y=100, relwidth=1, height=925)

        Label(self.purchaseF, text="Bill Details", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold"), width=85, pady=10).place(x=1, y=1, relwidth=1, height=60)

        # -------- Purchaser's Details Frame ---------
        self.purchaseDF = LabelFrame(self.purchaseF, text="Purchaser's Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"), padx=20, pady=5)
        self.purchaseDF.place(x=1, y=80, relwidth=1, height=150)

        # Purchaser's Name
        Label(self.purchaseDF, text="Purchaser's Name:", font=(
            "times new roman", 18, "bold"), pady=10).grid(row=0, column=0, pady=15, sticky=E)
        self.pname = StringVar()
        self.pnameC = ttk.Combobox(self.purchaseDF, width=30, font=(
            "arial", 18, "bold"), textvariable=self.pname, postcommand=self.updatePurchaserList)
        self.pnameC.grid(row=0, column=1,
                         padx=30, pady=15, sticky=W)

        # Purchaser's GST
        Label(self.purchaseDF, text="Purchaser's GST No.:", font=(
            "times new roman", 18, "bold"), pady=10).grid(row=0, column=2, pady=15, sticky=E)
        self.gst = Entry(self.purchaseDF, width=30, bd=3, font=(
            "arial", 18, "bold"))
        self.gst.grid(row=0, column=3,
                      padx=30, pady=15, sticky=W)

        # Add Purchaser Button
        Button(self.purchaseDF, text="Add Purchaser", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=self.addPurchaser).grid(row=0, column=4, padx=10, pady=15)

        # -------- Purchase Bill Details Frame ---------
        self.purchaseBF = LabelFrame(self.purchaseF, text="Purchaser's Details", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"), padx=100, pady=5)
        self.purchaseBF.place(x=1, y=250, relwidth=1, height=580)

        # Billing Year
        Label(self.purchaseBF, text="Billing Year:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=0, column=0, pady=15, sticky=E)
        self.byear = IntVar()
        self.byearE = Entry(self.purchaseBF, width=15, bd=3, font=(
            "arial", 22, "bold"), textvariable=self.byear)
        self.byearE.grid(row=0, column=1, padx=50, pady=15, sticky=W)

        # Billing Month
        Label(self.purchaseBF, text="Billing Month:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=0, column=2, pady=15, sticky=E)
        self.bmonth = IntVar()
        self.bmonthE = Entry(self.purchaseBF, width=15, bd=3, font=(
            "arial", 22, "bold"), textvariable=self.bmonth)
        self.bmonthE.grid(row=0, column=3, padx=50, pady=15, sticky=W)

        # Billing Day
        Label(self.purchaseBF, text="Billing Day:", font=(
            "times new roman", 22, "bold"), pady=10).grid(row=0, column=4, pady=15, sticky=E)
        self.bday = IntVar()
        self.bdayE = Entry(self.purchaseBF, width=15, bd=3, font=(
            "arial", 22, "bold"), textvariable=self.bday)
        self.bdayE.grid(row=0, column=5, padx=50, pady=15, sticky=W)

        # Taxes Frame
        self.taxF = LabelFrame(self.purchaseBF, text="Taxes", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"), padx=55, pady=5)
        self.taxF.place(x=1, y=100, width=1100, height=395)

        # Columns
        Label(self.taxF, text="", font=(
            "times new roman", 20, "bold"), pady=10, padx=15).grid(row=0, column=0, pady=10, padx=25)
        Label(self.taxF, text="Taxable Amount", font=(
            "times new roman", 20, "bold"), pady=10, padx=15).grid(row=0, column=1, pady=10, padx=25)
        Label(self.taxF, text="CGST", font=(
            "times new roman", 20, "bold"), pady=10, padx=15).grid(row=0, column=2, pady=10, padx=25)
        Label(self.taxF, text="SGST", font=(
            "times new roman", 20, "bold"), pady=10, padx=15).grid(row=0, column=3, pady=10, padx=25)
        Label(self.taxF, text="IGST", font=(
            "times new roman", 20, "bold"), pady=10, padx=15).grid(row=0, column=4, pady=10, padx=25)

        # Rows
        Label(self.taxF, text="5%", font=(
            "times new roman", 20, "bold"), pady=10).grid(row=1, column=0, pady=5)
        Label(self.taxF, text="12%", font=(
            "times new roman", 20, "bold"), pady=10).grid(row=2, column=0, pady=5)
        Label(self.taxF, text="18%", font=(
            "times new roman", 20, "bold"), pady=10).grid(row=3, column=0, pady=5)
        Label(self.taxF, text="28%", font=(
            "times new roman", 20, "bold"), pady=10).grid(row=4, column=0, pady=5)

        # GST 5%
        self.bamt_5 = DoubleVar()
        self.bamt_5E = Entry(self.taxF, width=15, bd=3, font=(
            "arial", 20, "bold"), textvariable=self.bamt_5)
        self.bamt_5E.grid(row=1, column=1, padx=25, pady=10)
        self.bgst_5 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_5).grid(row=1, column=2, padx=25, pady=10)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_5).grid(row=1, column=3, padx=25, pady=10)
        self.bigst_5 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_5).grid(row=1, column=4, padx=25, pady=10)
        self.bamt_5E.bind('<Return>', self.updateGST_5)

        # GST 12%
        self.bamt_12 = DoubleVar()
        self.bamt_12E = Entry(self.taxF, width=15, bd=3, font=(
            "arial", 20, "bold"), textvariable=self.bamt_12)
        self.bamt_12E.grid(row=2, column=1, padx=25, pady=10)
        self.bgst_12 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_12).grid(row=2, column=2, padx=25, pady=10)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_12).grid(row=2, column=3, padx=25, pady=10)
        self.bigst_12 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_12).grid(row=2, column=4, padx=25, pady=10)
        self.bamt_12E.bind('<Return>', self.updateGST_12)

        # GST 18%
        self.bamt_18 = DoubleVar()
        self.bamt_18E = Entry(self.taxF, width=15, bd=3, font=(
            "arial", 20, "bold"), textvariable=self.bamt_18)
        self.bamt_18E.grid(row=3, column=1, padx=25, pady=10)
        self.bgst_18 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_18).grid(row=3, column=2, padx=25, pady=10)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_18).grid(row=3, column=3, padx=25, pady=10)
        self.bigst_18 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_18).grid(row=3, column=4, padx=25, pady=10)
        self.bamt_18E.bind('<Return>', self.updateGST_18)

        # GST 28%
        self.bamt_28 = DoubleVar()
        self.bamt_28E = Entry(self.taxF, width=15, bd=3, font=(
            "arial", 20, "bold"), textvariable=self.bamt_28)
        self.bamt_28E.grid(row=4, column=1, padx=25, pady=10)
        self.bgst_28 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_28).grid(row=4, column=2, padx=25, pady=10)
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bgst_28).grid(row=4, column=3, padx=25, pady=10)
        self.bigst_28 = DoubleVar()
        Entry(self.taxF, width=10, bd=2, font=(
            "arial", 20, "bold"), textvariable=self.bigst_28).grid(row=4, column=4, padx=25, pady=10)
        self.bamt_28E.bind('<Return>', self.updateGST_28)

        # Total Amount
        Label(self.purchaseBF, text="Total Amount:", font=(
            "times new roman", 22, "bold"), pady=10).place(x=1260, y=100)
        self.btamt = IntVar()
        self.btamtE = Label(self.purchaseBF, width=15, bd=3, fg=navy, font=(
            "arial", 22, "bold"), textvariable=self.btamt)
        self.btamtE.place(x=1220, y=165)

        # Total Amount Button
        Button(self.purchaseBF, text="Total Amount", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=self.totalBill).place(x=1200, y=240)

        # Add Bill Button
        Button(self.purchaseBF, text="Add Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=self.createBill).place(x=1200, y=335)

        # Home Button
        add_client_btn = Button(self.purchaseBF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(x=1200, y=430)

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
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
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
                # print("The SQLite connection is closed")

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
        year = self.byear.get()
        if not year:
            messagebox.showerror(
                title="Error", message="Enter a Year!!")
            return
        month = self.bmonth.get()
        if not month:
            messagebox.showerror(
                title="Error", message="Enter a Month!!")
            return
        day = self.bday.get()
        if not day:
            messagebox.showerror(
                title="Error", message="Enter a Day!!")
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
        constraints = [year, month, day, tax_amt, bgst_5, bigst_5,
                       bgst_12, bigst_12, bgst_18, bigst_18, bgst_28, bigst_28, total_amt]
        self.insertBill(name, constraints)

    def insertBill(self, name, constraints):
        try:
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute(
                """SELECT MAX(pb_no) FROM purchase_bill WHERE pb_year=?;""", (constraints[0],))
            val = cursor.fetchall()
            if val[0][0] == None:
                bill_no = 1
            else:
                bill_no = val[0][0]+1
            cursor.execute(
                """SELECT p_id FROM purchaser WHERE p_name=?;""", (name,))
            val = cursor.fetchall()
            p_id = val[0][0]
            constraints.insert(0, p_id)
            constraints.insert(0, bill_no)
            cursor.execute(
                """INSERT INTO purchase_bill (pb_no, p_id, pb_year, pb_month, pb_day, pb_tax_amt, pb_gst_5, pb_igst_5, pb_gst_12, pb_igst_12, pb_gst_18, pb_igst_18, pb_gst_28, pb_igst_28, pb_total_amt)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);""", tuple(constraints))
            sqliteConnection.commit()
            messagebox.showinfo(title="Successful",
                                message="Bill added successfully with Bill No:"+str(bill_no)+"!!")
            cursor.close()

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to insert into purchase_bill table", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                # print("The SQLite connection is closed")

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
            "times new roman", 42, "bold"), pady=2).place(x=1, y=2, relwidth=1)

        # --------- Sales Options ---------
        self.purchaseF = LabelFrame(self, text="Purchases", bd=6, relief=GROOVE, labelanchor=N, font=(
            "times new roman", 30, "bold"), padx=220, pady=10)
        self.purchaseF.place(x=1, y=130, relwidth=1, height=895)

        Label(self.purchaseF, text="Check & Edit Bill Payment", bd=4, relief=GROOVE, font=(
            "times new roman", 26, "bold"), width=85, pady=10).place(x=20, y=20, relwidth=1, height=60)

        # --------- Check Bill Status Frame ------------
        self.selectBillF = LabelFrame(self.purchaseF, text="  Search Bill Payment  ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"), padx=20, pady=20)
        self.selectBillF.place(x=30, y=100, width=700, height=320)

        # Billing Year
        Label(self.selectBillF, text="Billing Year:", font=(
            "times new roman", 18, "bold"), pady=10).place(x=20, y=10)
        self.byear = Entry(self.selectBillF, width=10, bd=3, font=(
            "arial", 18, "bold"))
        self.byear.place(x=180, y=15)

        # Billing Month
        Label(self.selectBillF, text="Billing Month:", font=(
            "times new roman", 18, "bold"), pady=10).place(x=370, y=10)
        self.bmonth = Entry(self.selectBillF, width=5, bd=3, font=(
            "arial", 18, "bold"))
        self.bmonth.place(x=550, y=15)

        # Purchaser's Name
        Label(self.selectBillF, text="Purchaser's Name:", font=(
            "times new roman", 18, "bold"), pady=10).place(x=20, y=80)
        self.pname = StringVar()
        self.pnameC = ttk.Combobox(self.selectBillF, width=28, font=(
            "arial", 18, "bold"), textvariable=self.pname, postcommand=self.updatePurchaserList)
        self.pnameC.place(x=240, y=85)

        # Search Bills Button
        Button(self.selectBillF, text="Search Bill", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=30, font=(
            "arial", 18, "bold"), command=self.searchBill).place(x=120, y=150)

        # --------- Edit Bill Status Frame ------------
        self.editStatusF = LabelFrame(self.purchaseF, text="  Update Bill Payment  ", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 22, "bold"), padx=20, pady=20)
        self.editStatusF.place(x=780, y=100, width=700, height=320)

        # Bill No.
        Label(self.editStatusF, text="Billing No:", font=(
            "times new roman", 18, "bold")).grid(row=0, column=0, padx=10, pady=10, sticky=E)
        self.bno = StringVar()
        self.bnoC = Entry(self.editStatusF, width=30, bd=3, font=(
            "arial", 18, "bold"), textvariable=self.bno)
        self.bnoC.grid(row=0, column=1, padx=10, pady=10)

        # Payment Status
        Label(self.editStatusF, text="Payment Status:", font=(
            "times new roman", 18, "bold")).grid(row=1, column=0, padx=15, pady=20, sticky=E)
        self.estatus = Entry(self.editStatusF, width=30, bd=3, font=(
            "arial", 18, "bold"))
        self.estatus.grid(row=1, column=1, padx=10, pady=10)

        # Current Bill Details Button
        Button(self.editStatusF, text="Update Status", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", pady=10, width=30, font=(
            "arial", 18, "bold"), command=self.editStatus).place(x=120, y=150)

        # --------- View Bill Detials Frame ------------
        self.viewBillF = LabelFrame(self.purchaseF, text="View Bill Detials", bd=6, relief=GROOVE, labelanchor=NW, font=(
            "times new roman", 18, "bold"), padx=10, pady=10)
        self.viewBillF.place(x=30, y=440, width=1000, height=360)
        self.displayBillF = Frame(self.viewBillF)
        self.displayBillF.place(x=10, y=20, width=950, height=280)
        self.displayText = scrolledtext.ScrolledText(
            self.displayBillF, height=280, font=("Courier",
                                                 12, "bold"), padx=10, pady=10)
        self.displayText.insert(INSERT,
                                "\n\n\n\n\n\n\t\tSelect Billing Year or Purchaser's Name to see the entries!! ")
        self.displayText.configure(state='disabled')
        self.displayText.pack(side="left", fill="both", expand=True)
        Label(self.viewBillF, text="Bill No.\tYear\t\t            Purchaser's Name\t\t\tAmount\t              Payment", font=(
            "times new roman", 15, "bold")).place(x=30, y=0)

        # Home Button
        Button(self.purchaseF, text="Home", cursor="hand2", bd=5, relief=GROOVE, bg="cadetblue", width=25, pady=20, font=(
            "arial", 18, "bold"), command=lambda: controller.show_frame(Home)).place(x=1080, y=600)

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
            # constraints.append("pb_year")
            constraints.append(b_year)
            b_month = self.bmonth.get()
            if b_month:
                # constraints.append("pb_month")
                constraints.append(b_month)
        if not len(constraints):
            messagebox.showerror(
                title="Error", message="Select atleast 1 critera to search!!")
            return
        self.selectBill(constraints)

    def selectBill(self, constraints):
        try:
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            s = ""
            if constraints[0] == "p_name":
                cursor.execute(
                    "SELECT p_id FROM purchaser WHERE p_name =?;", (constraints[1],))
                rows = cursor.fetchall()
                p_id = rows[0][0]
                if len(constraints) == 2:
                    cursor.execute(
                        "SELECT pb_no, pb_year, pb_total_amt, pb_status FROM purchase_bill WHERE p_id=?;", (p_id,))
                elif len(constraints) == 3:
                    cursor.execute("SELECT pb_no, pb_year, pb_total_amt, pb_status FROM purchase_bill WHERE p_id=? AND pb_year=?;",
                                   (p_id, constraints[2],))
                else:
                    cursor.execute("SELECT pb_no, pb_year, pb_total_amt, pb_status FROM purchase_bill WHERE p_id=? AND pb_year=? AND pb_month=?;",
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
                    cursor.execute("SELECT pb_no, pb_year, p_id, pb_tax_amt, pb_status FROM purchase_bill WHERE pb_year=?;",
                                   (constraints[0],))
                else:
                    cursor.execute("SELECT pb_no, pb_year, p_id, pb_tax_amt, pb_status FROM purchase_bill WHERE pb_year=? AND pb_month=?;",
                                   (constraints[0], constraints[1]))
                rows = cursor.fetchall()
                for row in rows:
                    s += "\n"+str(row[0]).center(7)
                    s += str(row[1]).center(8)
                    cursor.execute(
                        "SELECT p_name FROM purchaser WHERE p_id =?;", (row[2],))
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
                # print("The SQLite connection is closed")

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
            sqliteConnection = sqlite3.connect('Billing.db')
            cursor = sqliteConnection.cursor()
            # print("Connected to SQLite")
            cursor.execute(
                "UPDATE purchase_bill SET pb_status=? WHERE pb_year=? AND pb_no=?;", constraints)
            messagebox.showinfo(
                title="Successfull", message="Payment Status updated Successfully!!")
            sqliteConnection.commit()
            cursor.close()
            self.selectBill([constraints[1]])

        except sqlite3.Error as error:
            messagebox.showerror(
                title="Failed to Get Data form Database", message=error)
        finally:
            if (sqliteConnection):
                sqliteConnection.close()
                # print("The SQLite connection is closed")

    def updatePurchaserList(self):
        getPurchaserList()
        self.pnameC['values'] = purchasers


def getCientList():
    sqliteConnection = sqlite3.connect('Billing.db')
    cursor = sqliteConnection.cursor()
    # print("Connected to SQLite")
    cursor.execute("""SELECT c_name FROM client;""")
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
    sqliteConnection = sqlite3.connect('Billing.db')
    cursor = sqliteConnection.cursor()
    # print("Connected to SQLite")
    cursor.execute("""SELECT p_name FROM purchaser;""")
    rows = cursor.fetchall()
    purchasers.clear()
    for i in range(len(rows)):
        purchasers.append(rows[i][0])
    purchasers.sort()
    sqliteConnection.commit()
    cursor.close()
    if (sqliteConnection):
        sqliteConnection.close()


if __name__ == "__main__":
    app = App()
    app.mainloop()
