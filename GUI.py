import tkinter as tk
import TestCaseUtility as tcu
from tkinter.ttk import *


class TestcaseUtilityMain(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.winfo_toplevel().title("TestCaseUtility")
        self.create_widget()
        self.open_widget()
        self.exit_widget()
        tk.Frame(root).config(highlightthickness=2, padx=2)
        root.geometry("250x116")
        root.option_add("*Button.Background", "black")
        root.option_add("*Button.Foreground", "green")
        root["bg"] = "#aaaaaa"

    def create_widget(self):
        create_button = Button(master=root, text="Create", style='TButton', command=self.create_file)
        create_button.pack(side="left")

    def open_widget(self):
        open_button = Button(master=root, text="Open", style='TButton', command=self.open_file)
        open_button.pack(side="right")

    def exit_widget(self):
        quit_button = Button(text="Exit", command=self.master.destroy)
        quit_button.pack(side="bottom")

    def create_file(self):
        tcu.create_file_function()

    def open_file(self):
        tcu.open_function()


root = tk.Tk()
app = TestcaseUtilityMain(master=root)
style = Style()
style.configure('TButton', font=('calibri', 10, 'bold'), borderwidth='4')
app.mainloop()
