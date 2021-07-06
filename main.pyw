import excel_writer
import filehandler
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import json
import os
import sys
import threading
import global_vars
try:
    from PIL import Image, ImageTk
except ModuleNotFoundError:
    print("pip3 install --user pillow")
    os._exit(0)


class Gui:
    def __init__(self):
        self.main = tk.Tk()
        self.running = False
        self.main.title("K-faktor script")
        self.main.geometry("800x800")
        self.folder = tk.StringVar()
        self.out_folder = tk.StringVar()
        self.out_folder.set(f"Output directory:   {global_vars.settings['out_folder']}")
        self.folder.set(f"Current folder:   {global_vars.settings['data_folder']}")
        self.progress_str = tk.StringVar()
        self.use_default = tk.BooleanVar()
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
        self.filebutton = tk.Button(self.main, text="Choose data folder", command=self.prompt_data_folder)
        self.startbutton = tk.Button(self.main, text="Start", command=self.start_thread, width=25, height=2, font=("Helvetica", 16))
        self.save_button = tk.Button(self.main, text="Save settings", command=self.save_folder)
        self.out_folder_button = tk.Button(self.main, text="Choose outfolder", command=self.prompt_out_folder)
        self.disp_out_folder = tk.Label(self.main, textvariable=self.out_folder, font=("Helvetica", 13))
        self.disp_folder_var = tk.Label(self.main, textvariable=self.folder, font=("Helvetica", 13))
        self.use_default_button = tk.Checkbutton(self.bottom, text="Use default folders", variable=self.use_default, onvalue=True, offvalue=False)
        self.clear_button = tk.Button(self.main, text="Clear settings", command=self.clear_settings)
        self.progress = ttk.Progressbar(self.main, orient=tk.HORIZONTAL, length=200, mode="determinate")
        self.progress["value"] = 0
        self.progress_label = tk.Label(self.main, textvariable=self.progress_str)

        # Pack 
        self.title.pack(in_=self.top, pady=(10,10))
        self.filebutton.pack(in_=self.top)
        self.disp_folder_var.pack(in_=self.top, pady=(5,5))
        self.out_folder_button.pack(in_=self.top)
        self.disp_out_folder.pack(in_=self.top)
        self.save_button.pack(in_=self.top, pady= (5,5))
        self.vim_label.pack(in_=self.top)
        self.startbutton.pack(in_=self.bottom) #, pady=(10,10))
        self.progress_label.pack(in_=self.bottom)#, pady=(10,10))
        self.progress.pack(in_=self.bottom, pady=(10,10))
        self.use_default_button.pack(in_=self.bottom, side=tk.RIGHT, pady=(0,5))
        self.clear_button.pack(in_=self.bottom, side=tk.LEFT, pady=(0,5))
        
    def prompt_data_folder(self):
        if self.running:
            messagebox.showwarning("Warning", "Cannot change datafolder while process is runnig.")
        else:
            tmp = filedialog.askdirectory().replace("\\", "/")
            if tmp != "":
                global_vars.settings["data_folder"] = tmp
                self.folder.set(f"Data folder: {global_vars.settings['data_folder']}")

    def prompt_out_folder(self):
        if self.running:
            messagebox.showwarning("Warning", "Cannot change outfolder while process is runnig.")
        else:
            global_vars.settings["out_folder"] = filedialog.askdirectory().replace("\\", "/")
            self.out_folder.set(f"Data folder: {global_vars.settings['out_folder']}")
    def save_folder(self):
        save_config()
    
    def clear_settings(self):
        global_vars.settings["data_folder"] = f"{os.getcwd()}/data".replace('\\', '/')
        global_vars.settings["out_folder"] = os.getcwd().replace('\\', '/')
        self.out_folder.set(f"Output directory:   {global_vars.settings['out_folder']}")
        self.folder.set(f"Current folder:   {global_vars.settings['data_folder']}")
       
    def start_thread(self):
        if self.running:
            messagebox.showwarning("Thread error", "Warning! Peocess already running")
            return None
        else:
            thread = threading.Thread(target=self.yeet)
            thread.start()
        self.running = True
         
    def yeet(self):
        print(self.use_default.get())
        if self.use_default.get():
            self.clear_settings()
        try:
            self.progress["value"] = 25
            self.progress_str.set("Renaming and moving files...")
            self.main.update_idletasks()
            filehandler.rename_and_move_files(global_vars.settings)
            self.progress["value"] = 50
            self.progress_str.set("Converting PDF files to CSV files...")
            self.main.update_idletasks()
            filehandler.convert_pdffiles_to_csv(global_vars.settings)
            self.progress["value"] = 99
            self.progress_str.set("Moving data to excel...")
            self.main.update_idletasks()
            excel_writer.data_to_excel(global_vars.settings)
            self.progress["value"] = 100
            self.progress_str.set("Done")
            self.main.update_idletasks()
            self.running = False
        except Exception as e:
            self.progress_str.set("Error")
            self.progress["value"] = 0
            self.running = False
            messagebox.showerror("Exception", e)

def parse_config():
    global_vars.settings = dict()
    if os.path.isfile("./config.json"):
        with open("./config.json", "r") as f:
            try:
                tmp = json.load(f)
            except json.JSONDecodeError:
                print("JSON decoding error, unable to read config file. Using default config...")
                global_vars.settings["data_folder"] = f"{os.getcwd()}/data".replace('\\', '/')
                global_vars.settings["out_folder"] = os.getcwd().replace('\\', '/')
                return None

            if "data_folder" in tmp:
                global_vars.settings["data_folder"] = tmp["data_folder"]
            else:
                global_vars.settings["data_folder"] = os.getcwd().replace("\\", "/") + "/data"
            if "out_folder" in tmp:
                global_vars.settings["out_folder"] = tmp["out_folder"]
            else:
                global_vars.settings["out_folder"] = os.getcwd().replace("\\", "/")
    else:
        global_vars.settings["data_folder"] = f"{os.getcwd()}/data".replace('\\', '/')
        global_vars.settings["out_folder"] = os.getcwd().replace('\\', '/')
        
def save_config():
    with open("./config.json","w") as f:
        json.dump(global_vars.settings, f, indent=4)

if __name__ == "__main__":
    print(sys.path[0])
    os.chdir(sys.path[0])
    parse_config() 
    if "PROMPT" in os.environ:
        filehandler.rename_and_move_files(global_vars.settings)
        filehandler.convert_pdffiles_to_csv(global_vars.settings)
        excel_writer.data_to_excel(global_vars.settings)
    else:
        gui = Gui()
     