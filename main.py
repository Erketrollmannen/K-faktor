import excel_writer
import filehandler
import tkinter as tk
from tkinter import filedialog
import config

class Gui:
    def __init__(self):
        self.main = tk.Tk()
        self.data_folder = "./data"
        self.filebutton = tk.Button(self.main, text="Choose data folder", command=self.prompt_file)
        self.startbutton = tk.Button(self.main, text="Start", command=self.yeet)
        self.filebutton.grid(row=0, column=0)
        self.startbutton.grid(row=1, column=1)
        
        tk.mainloop()

    def prompt_file(self):
        self.data_folder = filedialog.askdirectory()
       
    def yeet(self):
        filehandler.rename_and_move_files()
        filehandler.convert_pdffiles_to_csv()
        excel_writer.data_to_excel()





if __name__ == "__main__":
    # gui = Gui()
    filehandler.rename_and_move_files()
    filehandler.convert_pdffiles_to_csv()
    excel_writer.data_to_excel()