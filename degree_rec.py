import datetime
import time
import tkinter as tk
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
from datetime import datetime
import random

file1 = ''



window = tk.Tk()
window.geometry("250x450")
window.title('Deg Rec')

buton_file_1 = tk.Button(
    text="Browse",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="green",
    fg="black",
)
buton_file_1.pack()

buton_deg_reg = tk.Button(
    text="Deg reg",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="green",
    fg="black",
)
buton_deg_reg.pack()

def file_1(event):
    global file1
    file1 = askopenfilename(
        filetypes=[('text', '*.xlsx'), ('all files', '*.*')]
    )


def deg_reg(event):

    global file1
    wb = load_workbook(file1)
    # sheet = wb.active

    # load and write:
    wb = load_workbook(file1)
    sheet = wb['Sheet']
    degree='22°'
    numero_cell=['A' + str(i) for i in range(1, 22)]
    degree_cell=['B' + str(i) for i in range(1, 22)]
    x=1
    for i in numero_cell[1:21]:
        if sheet[i].value is not None:
            sheet[i].value=x
            x+=1
            sheet[degree_cell[x]].value=degree
            break
    wb.save(file1)

buton_file_1.bind("<Button-1>", file_1)
buton_deg_reg.bind("<Button-1>", deg_reg)
window.mainloop()