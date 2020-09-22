import tkinter as tk
import TestCaseUtility as tcu


class TestcaseUtilityMain(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.winfo_toplevel().title("TestCaseUtility")
        self.create_widget()
        self.open_widget()
        self.exit_widget()
        root.geometry("250x116")
        root.option_add("*Button.Background", "black")
        root.option_add("*Button.Foreground", "green")
        root["bg"] = "#aaaaaa"

    def create_widget(self):
        create_button = tk.Button(master=root, text="Create", command=self.create_file)
        create_button.pack(side="left")

    def open_widget(self):
        open_button = tk.Button(master=root, text="Open", command=self.open_file)
        open_button.pack(side="right")

    def exit_widget(self):
        quit_button = tk.Button(text="QUIT", fg="red", command=self.master.destroy)
        quit_button.pack(side="bottom")

    def create_file(self):
        tcu.create_file_function()

    def open_file(self):
        tcu.open_function()


root = tk.Tk()
app = TestcaseUtilityMain(master=root)
app.mainloop()
