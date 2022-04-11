import win32com.client
import tkinter as Tk
from tkinter import filedialog
import sys

root = Tk.Tk()
root.withdraw()

filename = filedialog.askopenfilename()

if filename == "":
    print("nie wybrano pliku! zamykam program..")
    sys.exit()

word = win32com.client.Dispatch("Word.Application")
word.visible = 1
# set the visible to 0, if you dont want to see the word application

wordObj = word.Documents.Open(filename)
wordObj.SaveAs(filename, FileFormat=16)
# File format 16 refers to word file