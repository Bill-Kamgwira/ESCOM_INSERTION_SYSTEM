from tkinter import *

from tkinter import ttk

import datetime

import tkinter.messagebox

import os

import openpyxl


class Inventory:
    def __init__(self, root):

        self.root = root

        self.root.title("ESCOM Inventory Insertion Tool")

        self.root.configure(background="green")

        # Main Frames config

        MainFrame = Frame(self.root, bd=10, bg="green", relief=RIDGE)
        MainFrame.pack()

        LeftFrame = LabelFrame(
            MainFrame,
            text="Information Panel",
            font=("arial", 15, "bold"),
            bd=10,
            bg="floral white",
            padx=1,
            relief=RIDGE,
        )
        LeftFrame.grid(row=0, column=0)

        RightFrame = LabelFrame(
            MainFrame,
            text="Review Panel",
            font=("arial", 15, "bold"),
            bd=10,
            width=560,
            height=650,
            bg="floral white",
            padx=1,
            relief=RIDGE,
        )
        RightFrame.grid(row=1, column=0)

        # Frames for the Text, Label and Entry widget

        LeftFrame0 = Frame(
            LeftFrame,
            bd=5,
            width=855,
            height=145,
            bg="floral white",
            padx=5,
            relief=RIDGE,
        )
        LeftFrame0.grid(
            row=0,
            column=0,
        )

        LeftFrame1 = Frame(
            LeftFrame,
            bd=5,
            width=855,
            height=170,
            bg="floral white",
            padx=5,
            relief=RIDGE,
        )
        LeftFrame1.grid(
            row=1,
            column=0,
        )

        RightFrame1 = Frame(
            RightFrame,
            bd=5,
            width=450,
            height=560,
            bg="floral white",
            relief=RIDGE,
        )
        RightFrame1.grid(
            row=1,
            column=0,
        )

        RightFrame2 = Frame(
            RightFrame,
            bd=5,
            width=535,
            height=165,
            bg="floral white",
            relief=RIDGE,
        )
        RightFrame2.grid(
            row=2,
            column=0,
        )

        AcctOpen = StringVar()
        AppDate = StringVar()
        NCReR = StringVar()
        LCReR = StringVar()
        DateRev = StringVar()
        ProdType = StringVar()
        NoDays = StringVar()
        ProdCode = StringVar()
        Supplier = StringVar()
        Name = StringVar()
        Model = StringVar()
        SetDue = StringVar()
        Make = StringVar()

        # LeftFrame0 Content

        # FIRST PIECE

        # CONFIGURATION OF MOBILE FUNCTION

        def MobileOpt(evt):

            values = str(self.cboProdType.get())

            N_A = values

            today_date = datetime.date.today()

            if N_A == "Mobile":

                fill = "N/A"

                TAG.set(fill)

                NoDays.set(today_date)

            else:
                TAG.set("")

                NoDays.set(today_date)
            return

        self.lblProdType = Label(
            LeftFrame0,
            font=("arial", 15, "bold"),
            text="Product Type:",
            pady=2,
            bg="floral white",
        )
        self.lblProdType.grid(row=0, column=0, sticky="news")

        self.cboProdType = ttk.Combobox(
            LeftFrame0,
            textvariable=ProdType,
            state="readonly",
            font=("arial", 15, "bold"),
            width=15,
        )

        self.cboProdType.bind("<<ComboboxSelected>>", MobileOpt)
        self.cboProdType["value"] = ("", "Mobile", "PC", "Printer", "Other")
        self.cboProdType.current(0)
        self.cboProdType.grid(row=0, column=1)

        # SECOND PIECE
        self.lblProdCode = Label(
            LeftFrame0,
            font=("arial", 15, "bold"),
            text="Product Code:",
            pady=2,
            bg="floral white",
        )
        self.lblProdCode.grid(row=1, column=0, sticky="news")

        self.txtProdCode = Entry(
            LeftFrame0,
            textvariable=ProdCode,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=30,
            justify=LEFT,
        ).grid(row=1, column=1)

        # THIRD PIECE
        self.lblSupplier = Label(
            LeftFrame0,
            font=("arial", 15, "bold"),
            text="Supplier:",
            pady=2,
            bg="floral white",
        )
        self.lblSupplier.grid(row=1, column=2, sticky="news")

        self.txtSupplier = Entry(
            LeftFrame0,
            textvariable=Supplier,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=25,
            justify=LEFT,
        ).grid(row=1, column=3)

        # LeftFrame1 Content

        # FIRST PIECE
        self.lblName = Label(
            LeftFrame1,
            font=("arial", 15, "bold"),
            text="Name:",
            pady=2,
            bg="floral white",
        )
        self.lblName.grid(row=0, column=0, sticky="news")

        self.txtName = Entry(
            LeftFrame1,
            textvariable=Name,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=30,
            justify=LEFT,
        ).grid(row=0, column=1)

        # SECOND PIECE

        self.lblDate = Label(
            LeftFrame1,
            font=("arial", 15, "bold"),
            text="Date:",
            pady=2,
            bg="floral white",
        )
        self.lblDate.grid(row=0, column=2, sticky="news")

        self.txtDate = Entry(
            LeftFrame1,
            textvariable=NoDays,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=17,
            justify=LEFT,
        ).grid(row=0, column=3)

        # THIRD PIECE
        self.lblModel = Label(
            LeftFrame1,
            font=("arial", 15, "bold"),
            text="Model:",
            pady=2,
            bg="floral white",
        )
        self.lblModel.grid(row=1, column=0, sticky="news")

        self.txtModel = Entry(
            LeftFrame1,
            textvariable=Model,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=30,
            justify=LEFT,
        ).grid(row=1, column=1)

        # FOURTH PIECE
        self.lblMake = Label(
            LeftFrame1,
            font=("arial", 15, "bold"),
            text="Make:",
            pady=2,
            bg="floral white",
        )
        self.lblMake.grid(row=1, column=2, sticky="news")

        self.txtMake = Entry(
            LeftFrame1,
            textvariable=Make,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=17,
            justify=LEFT,
        ).grid(row=1, column=3)

        # FIFTH PIECE
        self.lblTag = Label(
            LeftFrame1,
            font=("arial", 15, "bold"),
            text="Tag:",
            pady=2,
            bg="floral white",
        )
        self.lblTag.grid(row=3, column=0, sticky="news")

        TAG = StringVar()

        self.txtTag = Entry(
            LeftFrame1,
            textvariable=TAG,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=30,
            justify=LEFT,
        ).grid(row=3, column=1)

        # SIXTH PIECE
        self.lblUser = Label(
            LeftFrame1,
            font=("arial", 15, "bold"),
            text="User:",
            pady=2,
            bg="floral white",
        )
        self.lblUser.grid(row=2, column=0, sticky="news")

        USER = StringVar()

        self.txtUser = Entry(
            LeftFrame1,
            textvariable=USER,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=30,
            justify=LEFT,
        ).grid(row=2, column=1)

        # SEVENTH PIECE
        self.lblStation = Label(
            LeftFrame1,
            font=("arial", 15, "bold"),
            text="Station:",
            pady=2,
            bg="floral white",
        )
        self.lblStation.grid(row=2, column=2, sticky="news")

        Station = StringVar()

        self.txtStation = Entry(
            LeftFrame1,
            textvariable=Station,
            font=("arial", 15, "bold"),
            bd=8,
            fg="black",
            width=17,
            justify=LEFT,
        ).grid(row=2, column=3)

        # Content inside RightFrame1

        self.txtReceipt = Text(
            RightFrame1,
            height=18,
            width=71,
            font=(
                "arial",
                9,
                "bold",
            ),
        )
        self.txtReceipt.grid(row=0, column=0, sticky="news")

        # Buttons for RightFrame2

        # CONFIGURATION FOR ALL FRAMES

        for widget in LeftFrame.winfo_children():

            widget.grid_configure(padx=10, pady=5)

        # Configuration for Save to File Button
        def SaveToFile():
            # EXCEL EXPORTION CODE

            prodtype = ProdType.get()
            date = NoDays.get()
            prodcode = ProdCode.get()
            supplier = Supplier.get()
            name = Name.get()
            model = Model.get()
            make = Make.get()
            tag = TAG.get()
            user = USER.get()
            station = Station.get()

            filepath = "data.xlsx"

            if not os.path.exists(filepath):

                Info = openpyxl.Workbook()
                Sheet = Info.active
                Headings = [
                    "Product Type",
                    "Product Code",
                    "Name of Supplier",
                    "Name of Product",
                    "Model of Product",
                    "Make of Product",
                    "Name of User",
                    "Name of Station",
                    "TAG NUMBER",
                    "Date Of Record",
                ]

                Sheet.append(Headings)
                Info.save(filepath)

            Info = openpyxl.load_workbook(filepath)
            Sheet = Info.active

            NewRow = [
                prodtype,
                prodcode,
                supplier,
                name,
                model,
                make,
                user,
                station,
                tag,
                date,
            ]
            Sheet.append(NewRow)

            Info.save(filepath)

            iconfirm = tkinter.messagebox.showinfo("Saved!", "INFORMATION SAVED!")

            return iconfirm

        self.btnSave = Button(
            RightFrame2,
            padx=18,
            pady=2,
            bd=4,
            fg="black",
            font=("arial", 9, "bold"),
            width=9,
            bg="floral white",
            text="Save",
            command=SaveToFile,
        ).grid(row=0, column=0)

        # Configuration for the Review Button

        def Review():
            self.txtReceipt.insert(END, "\nProduct Type:\t\t" + ProdType.get() + "\n")
            self.txtReceipt.insert(END, "\nProduct Code:\t\t" + ProdCode.get() + "\n")
            self.txtReceipt.insert(
                END, "\nName of Supplier:\t\t" + Supplier.get() + "\n"
            )
            self.txtReceipt.insert(END, "\nName of Product:\t\t" + Name.get() + "\n")
            self.txtReceipt.insert(END, "\nModel of Product:\t\t" + Model.get() + "\n")
            self.txtReceipt.insert(END, "\nMake of Prodcut:\t\t" + Make.get() + "\n")
            self.txtReceipt.insert(END, "\nName of User:\t\t" + USER.get() + "\n")
            self.txtReceipt.insert(END, "\nName of Station:\t\t" + Station.get() + "\n")
            self.txtReceipt.insert(END, "\nTAG NUMBER:\t\t" + TAG.get() + "\n")
            self.txtReceipt.insert(END, "\nDate Of Record:\t\t" + NoDays.get() + "\n")

        self.btnReview = Button(
            RightFrame2,
            padx=18,
            pady=2,
            bd=4,
            fg="black",
            font=("arial", 9, "bold"),
            width=9,
            bg="floral white",
            text="Review",
            command=Review,
        ).grid(row=0, column=1)

        # RESET BUTTON CONFIGURATION
        def Reset():
            ProdType.set("")
            NoDays.set("")
            ProdCode.set("")
            Supplier.set("")
            Name.set("")
            Model.set("")
            Make.set("")
            TAG.set("")
            USER.set("")
            Station.set("")
            self.txtReceipt.delete("1.0", END)

            return

        self.btnReset = Button(
            RightFrame2,
            padx=18,
            pady=2,
            bd=4,
            fg="black",
            font=("arial", 9, "bold"),
            width=9,
            bg="floral white",
            text="Reset",
            command=Reset,
        ).grid(row=0, column=2)

        # CONFIGURATION FOR THE EXIT BUTTON

        def iExit():

            iExit = tkinter.messagebox.askyesno(
                "Inventory Systems", "Confirm if you want to exit"
            )

            if iExit > 0:
                root.destroy()

            return

        self.btnExit = Button(
            RightFrame2,
            padx=18,
            pady=2,
            bd=4,
            fg="black",
            font=("arial", 9, "bold"),
            width=9,
            bg="floral white",
            text="Exit",
            command=iExit,
        ).grid(row=0, column=3)


# Execution code
if __name__ == "__main__":
    root = tkinter.Tk()

    application = Inventory(root)

    root.mainloop()
