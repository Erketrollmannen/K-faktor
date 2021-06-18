#! c:\Python
import excel_writer
import filehandler
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk
import json
import os
import sys
import threading
import global_vars

class Gui:
    def __init__(self):
        
        self.main = tk.Tk()
        self.running = False
        self.main.title("K-faktor script")
        self.main.geometry("800x800")
        self.folder = tk.StringVar()
        self.folder.set(f"Current folder:   {global_vars.settings['data_folder']}")
        self.progress_str = tk.StringVar()
        self.top = tk.Frame(self.main)
        self.bottom = tk.Frame(self.main)
        self.top.pack(side=tk.TOP)
        self.bottom.pack(side=tk.BOTTOM)
        self.create_widgets()
        tk.mainloop()

    def create_widgets(self):
        self.vim_neger = Image.open(os.getcwd().replace("\\", "/") + "./img/vim.png")
        self.img = ImageTk.PhotoImage(self.vim_neger)
        self.vim_label = tk.Label(image=self.img)
        self.title = tk.Label(self.main, text="K-faktor script", font=("Helvetica", 20))
        self.filebutton = tk.Button(self.main, text="Choose folder", command=self.prompt_file)
        self.startbutton = tk.Button(self.main, text="Start", command=self.start_thread, width=25, height=4, font=("Helvetica", 16))
        self.save_button = tk.Button(self.main, text="Save data folder", command=self.save_data_folder)
        self.disp_folder_var = tk.Label(self.main, textvariable=self.folder, font=("Helvetica", 14))
        self.progress = ttk.Progressbar(self.main, orient=tk.HORIZONTAL, length=200, mode="determinate")
        self.progress["value"] = 0
        self.progress_label = tk.Label(self.main, textvariable=self.progress_str)

        # Pack 
        self.title.pack(in_=self.top, pady=(15,15))
        self.filebutton.pack(in_=self.top)
        self.disp_folder_var.pack(in_=self.top, pady=(15,15))
        self.save_button.pack(in_=self.top, pady= (10,10))
        self.vim_label.pack(in_=self.top)
        self.startbutton.pack(in_=self.bottom, pady=(20,10))
        self.progress_label.pack(in_=self.bottom, pady=(10,10))
        self.progress.pack(in_=self.bottom, pady=(10,10))
        
    def prompt_file(self):
        if self.running:
            messagebox.showwarning("Warning", "Cannot change datafolder while process is runnig.")
        else:
            global_vars.settings["data_folder"] = filedialog.askdirectory().replace("\\", "/")
            self.folder.set(f"Data folder: {global_vars.settings['data_folder']}")

    def save_data_folder(self):
        save_config()
       
    def start_thread(self):
        if self.running:
            messagebox.showwarning("Thread error", "Warning! Peocess already running")
            return None
        else:
            thread = threading.Thread(target=self.yeet)
            thread.start()
        self.running = True
         
    def yeet(self):
        self.progress["value"] = 25
        self.progress_str.set("Renaming and moving files...")
        self.main.update_idletasks()
        filehandler.rename_and_move_files()
        self.progress["value"] = 50
        self.progress_str.set("Converting PDF files to CSV files...")
        self.main.update_idletasks()
        filehandler.convert_pdffiles_to_csv()
        self.progress["value"] = 99
        self.progress_str.set("Moving data to excel...")
        self.main.update_idletasks()
        excel_writer.data_to_excel()
        self.progress["value"] = 100
        self.progress_str.set("Done")
        self.main.update_idletasks()
        self.running = False

def parse_config():
    if os.path.isfile("./config.json"):
        with open("./config.json", "r") as f:
            tmp = json.load(f)
            if "data_folder" in tmp:
                global_vars.settings = tmp
            else:
                global_vars.settings["data_folder"] = os.getcwd().replace("\\", "/") + "/data"
    else:
        global_vars.settings["data_folder"] = f"{os.getcwd()}/data".replace('\\', '/')

def save_config():
    with open("./config.json","w") as f:
        json.dump(global_vars.settings, f, indent=4)

if __name__ == "__main__":
    print(sys.path[0])
    os.chdir(sys.path[0])
    parse_config() 
    if "PROMPT" in os.environ:
        filehandler.rename_and_move_files()
        filehandler.convert_pdffiles_to_csv()
        excel_writer.data_to_excel()
    else:
        gui = Gui()
     