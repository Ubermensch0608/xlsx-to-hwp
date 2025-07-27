import tkinter as tk 
from tkinter import filedialog

def choose_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()  # GUI 창 숨김
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)

    return file_path 
