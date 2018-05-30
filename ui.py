import tkinter as tk

from datetime import datetime as dt
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter import ttk
from staffhours import StaffHours as sh

class CreateUi(ttk.Frame):
    def __init__(self, parent):
        # initialize the ui
        tk.Frame.__init__(self, parent)
        self.mainframe = parent

        # store the file paths
        self.services_chart_path = tk.StringVar()
        self.report_path = tk.StringVar()
        self.start_year = tk.IntVar()
        self.start_month = tk.IntVar()
        self.start_day = tk.IntVar()
        self.end_year = tk.IntVar()
        self.end_month = tk.IntVar()
        self.end_day = tk.IntVar()
        self.eha = tk.StringVar()

        # set the defalt file paths
        self.services_chart_path.set("\\\\tproserver\IT\Documentation\\1 - CaseManager\Service Codes v5.1 - Time Codes.xlsx")

        # create the labels, buttons, and path fields for the file paths
        ttk.Label(self.mainframe, text="Services Chart:").grid(column=1, row=1, sticky="W")
        ttk.Label(self.mainframe, text="Report:").grid(column=1, row=2, sticky="W")
        ttk.Entry(self.mainframe, textvariable=self.services_chart_path, width=100).grid(column=2, row=1, sticky="E")
        ttk.Entry(self.mainframe, textvariable=self.report_path, width=100).grid(column=2, row=2, sticky="E")
        ttk.Button(self.mainframe, text="Open", command=self.open_chart).grid(column=3, row=1, sticky="W")
        ttk.Button(self.mainframe, text="Open", command=self.open_report).grid(column=3, row=2, sticky="W")

        # create 6 text entry fields and 6 lables for capturing the period start and end dates
        ttk.Label(self.mainframe, text="Period Start Date:").grid(column=5, row=1, sticky="E")
        ttk.Label(self.mainframe, text="Period End Date:").grid(column=5, row=2, sticky="E")
        ttk.Label(self.mainframe, text="/").grid(column=7, row=1, sticky="E")
        ttk.Label(self.mainframe, text="/").grid(column=9, row=1, sticky="E")
        ttk.Label(self.mainframe, text="/").grid(column=7, row=2, sticky="E")
        ttk.Label(self.mainframe, text="/").grid(column=9, row=2, sticky="E")
        ttk.Entry(self.mainframe, textvariable=self.start_year, width=10).grid(column=10, row=1, sticky="E")
        ttk.Entry(self.mainframe, textvariable=self.start_month, width=3).grid(column=6, row=1, sticky="E")
        ttk.Entry(self.mainframe, textvariable=self.start_day, width=3).grid(column=8, row=1, sticky="E")
        ttk.Entry(self.mainframe, textvariable=self.end_year, width=10).grid(column=10, row=2, sticky="E")
        ttk.Entry(self.mainframe, textvariable=self.end_month, width=3).grid(column=6, row=2, sticky="E")
        ttk.Entry(self.mainframe, textvariable=self.end_day, width=3).grid(column=8, row=2, sticky="E")

        # create a dropdown menu for selecting if a report is an EAH report or not
        self.eha_list = [True, False]
        ttk.Combobox(self.mainframe, textvariable=self.eha, values=self.eha_list).grid(column=4, row=2)
        # create a process button and link it to the process method
        ttk.Button(self.mainframe, text="Process", command=self.process).grid(column=3, row=3, sticky="W")

    def open_chart(self):
        file = askopenfilename(title="Open the Services Chart")
        self.services_chart_path.set(file)

    def open_report(self):
        file = askopenfilename(title="Open the Staff Hours Report")
        self.report_path.set(file)

    def process(self):
        # create datetime objects out of the start and end IntVar variables
        start = dt(year=self.start_year.get(), month=self.start_month.get(), day=self.start_day.get())
        end = dt(year=self.end_year.get(), month=self.end_month.get(), day=self.end_day.get())

        # call the staff hours class and run its process method
        run = sh(
            self.report_path.get(),
            self.services_chart_path.get(),
            start,
            end,
            self.eha
        )
        run.process()

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Staff Hours Reporting")
    CreateUi(root)
    root.mainloop()
