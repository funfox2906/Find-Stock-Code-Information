import tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
root= tk.Tk()
root.title("Thong tin bang ma co phieu")
canvas1 = tk.Canvas(root, width = 200, height = 100, bg = 'lightsteelblue')
canvas1.pack()
def getExcel ():
    global df
    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel (import_file_path)

def exportExcel (out):
    global df
    df = pd.DataFrame(out)
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    df.to_excel (export_file_path, index = False, header=True)

def FindFromExcel(find):
    ind = df.index[df['Name'] == find].tolist()  # FIND BY NAME
    if not ind:
        messagebox.showinfo("Error", "Cannot find "+find)
    else:        
        messagebox.showinfo("Found it !", "Please choose file to save its content")
        exportExcel(pd.DataFrame(df.loc[ind]))
        

def retrieve_input(): 
    inputValue=textBox.get("1.0","end-1c")
    FindFromExcel(inputValue)
label = ttk.Label(root, text = "Enter Code")
label.place(x=15,y=103)
textBox=Text(root, height=1, width=5)
textBox.pack()
buttonCommit=Button(root, height=1, width=10, text="Find", 
                    command=lambda: retrieve_input())
buttonCommit.pack()
browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(100, 50, window=browseButton_Excel)
root.mainloop()