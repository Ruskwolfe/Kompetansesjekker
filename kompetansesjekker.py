import pandas as pd
from tkinter import Tk, Label, Button, filedialog, Entry, Text, END

class ExcelComparator:
    def __init__(self, master):
        self.master = master
        master.title("Excel Comparator")

        self.label = Label(master, text="Select the Excel files and enter the column name to compare.")
        self.label.pack()

        self.select_file1_button = Button(master, text="Select File 1", command=self.select_file1)
        self.select_file1_button.pack()

        self.select_file2_button = Button(master, text="Select File 2", command=self.select_file2)
        self.select_file2_button.pack()

        self.column_entry_label = Label(master, text="Enter Column Name:")
        self.column_entry_label.pack()

        self.column_name_entry = Entry(master)
        self.column_name_entry.pack()

        self.compare_button = Button(master, text="Compare", command=self.compare_files)
        self.compare_button.pack()

        self.result_text = Text(master, height=10, width=50)
        self.result_text.pack()

        self.file1 = None
        self.file2 = None

    def select_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    def select_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    def compare_files(self):
        column_name = self.column_name_entry.get()
        if self.file1 and self.file2 and column_name:
            df1 = pd.read_excel(self.file1)
            df2 = pd.read_excel(self.file2)
            # Find duplicates in the specified column
            duplicates = df1[df1[column_name].isin(df2[column_name])][column_name].unique()
            self.display_results(duplicates, column_name)
        else:
            self.result_text.insert(END, "Please select both files and specify the column name.\n")

    def display_results(self, duplicates, column_name):
        self.result_text.delete(1.0, END)
        if len(duplicates) > 0:
            result = "\n".join(duplicates.astype(str))
            self.result_text.insert(END, f"Duplicate values in '{column_name}':\n{result}")
        else:
            self.result_text.insert(END, "No duplicates found.")


root = Tk()
my_gui = ExcelComparator(root)
root.mainloop()
