import tkinter as tk
from openpyxl import load_workbook
from excel_project import execution, file_1, file_2, reset_data, date_recorder, deg_reg

file3 = 'date_hour.xlsx'

wb = load_workbook(file3)
sheet = wb.active

execution()

# grafica
window = tk.Tk()
window.geometry("250x550")
window.title('Excell app')

buton_file_1 = tk.Button(
    text="Browse",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg='green',
    fg="black",
)
buton_file_1.pack()

buton_file_2 = tk.Button(
    text="Report file",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="green",
    fg="black",
)
buton_file_2.pack()

buton_execution = tk.Button(
    text="Execute",
    width=15,
    height=2,
    font=("Arial", 15, "bold"),
    bg="green",
    fg="black",
)
buton_execution.pack()

buton_date_recorder = tk.Button(
    text="Date Recorder",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#3283a8",
    fg="black",
)
buton_date_recorder.pack()

buton_reset_data = tk.Button(
    text="Reset Data",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="yellow",
    fg="black",
)
buton_reset_data.pack()

file_1()

file_2()

reset_data()

date_recorder()

deg_reg()

buton_file_1.bind("<Button-1>", file_1)
buton_file_2.bind("<Button-1>", file_2)
buton_execution.bind("<Button-1>", execution)
buton_reset_data.bind("<Button-1>", reset_data)
buton_date_recorder.bind("<Button-1>", date_recorder)


window.mainloop()
