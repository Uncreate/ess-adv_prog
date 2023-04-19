import os
import shutil
import tkinter as tk
from win32com.client import Dispatch

__version__ = "1.0.0"

class FolderCreatorApp(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        master.title("Basic Part Creator")
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.folder_label = tk.Label(self, text="Enter a program number:")
        self.folder_label.grid(row=0,column=0, padx=5, pady=5,sticky="E")

        self.folder_entry = tk.Entry(self)
        self.folder_entry.grid(row=0,column=1, padx=5, pady=5, sticky="W")

        self.create_button = tk.Button(self)
        self.create_button["text"] = "Create Folder"
        self.create_button["command"] = self.create_folder
        self.create_button.grid(row=1,column=0, padx=5, pady=5)

        self.quit_button = tk.Button(
            self, text="Quit", command=self.master.destroy)
        self.quit_button.grid(row=1,column=1, padx=5, pady=5)

        self.message_text = tk.Text(self, height=10)
        self.message_text.grid(row=2,column=0, columnspan=2, padx=5, pady=5)

    def create_folder(self):
        try:
            folder_name = str(self.folder_entry.get()).zfill(6)
            folder_path = r"C:\Users\Public\PROGRAMING\CAM-1\{}".format(
                folder_name)
            if not os.path.exists(folder_path):
                self.copy_template(folder_path)
                self.create_shortcut(folder_path, folder_name)
                self.add_message(f"Folder {folder_path} created successfully!")
            else:
                self.add_message(f"Folder {folder_path} already exists!")
        except ValueError:
            self.add_message("Please enter a numeric value!")
        except Exception as e:
            self.add_message(f"Error creating folder: {e}")

    def copy_template(self, folder_path):
        self.add_message("Copying template to folder...")
        template_path = r"C:\Users\Public\PROGRAMING\CAM-1\TEMPLATES\Basic Program Folder\03xxxxx"
        shutil.copytree(template_path, folder_path)
        self.add_message(f"Template copied to {folder_path} successfully!")

    def create_shortcut(self, folder_path, folder_name):
        self.add_message("Creating shortcuts...")
        prz_folder_path = os.path.join(folder_path, "prz")
        os.makedirs(prz_folder_path, exist_ok=True)
        lnk_path = os.path.join(prz_folder_path, "200_PDM.lnk")
        target_path = r"C:\Live_Production_Vault\Projects\200-{}-01".format(
            folder_name)
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(lnk_path)
        shortcut.Targetpath = target_path
        shortcut.save()
        self.add_message(f"Shortcut created at {lnk_path} successfully!")

    def add_message(self, message):
        self.message_text.insert(tk.END, message + "\n")


root = tk.Tk()
app = FolderCreatorApp(master=root)
app.mainloop()
