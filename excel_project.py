"""
Excel automation project
"""
import time
import tkinter as tk
import random
from tkinter.filedialog import askopenfilename
from datetime import datetime
import openpyxl
from tkinter import *
import os
import sys
import subprocess

# GUI
WINDOW = Tk()
WINDOW.title('AppX v1.0')
WINDOW.geometry("1400x650")
WINDOW.configure(background="#1D6F42")
background_text = tk.Label(WINDOW,
                           text='KEEP\nCALM\nIT' + "'" + 'S JUST AN\nEXCEL\nFILE',
                           bg='#1D6F42',
                           font=("Arial", 35, "bold"),
                           fg="white")
background_text.place(x=20, y=165)

BUTTON_FILE_1 = tk.Button(
    text="Load",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg='#abf77e',
    fg="black")
BUTTON_FILE_1.place(x=450, y=20)
label1 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label1.place(x=700, y=20, width=600, height=85)
label1['text'] = 'No file selected! Click "Load" button to select a file.'


BUTTON_FILE_2 = tk.Button(
    text="Report",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#abf77e",
    fg="black")
BUTTON_FILE_2.place(x=550, y=120)
label2 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label2.place(x=800, y=120, width=600, height=85)
label2['text'] = 'No file selected! Click "Report" button to select a file.'


BUTTON_DATE_RECORDER = tk.Button(
    text="Rec",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#abf77e",
    fg="black")
BUTTON_DATE_RECORDER.place(x=450, y=220)
label3 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label3.place(x=650, y=220, width=600, height=85)
label3['text'] = 'Click "Rec" button to record data.'


BUTTON_EXECUTION = tk.Button(
    text="Execute",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#3EB489",
    fg="black")
BUTTON_EXECUTION.place(x=550, y=320)
label4 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label4.place(x=750, y=320, width=600, height=85)
label4['text'] = 'Click "Execute" button to create data report.'


BUTTON_RESET_DATA = tk.Button(
    text="Delete",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#E00201",
    fg="black")
BUTTON_RESET_DATA.place(x=450, y=420)
label5 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label5.place(x=650, y=420, width=600, height=85)
label5['text'] = 'Click "Delete" button to reset data.'

BUTTON_OPEN = tk.Button(
    text="Open Report",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="yellow",
    fg="black", )
BUTTON_OPEN.place(x=550, y=520)
label6 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label6.place(x=750, y=520, width=600, height=85)
label6['text'] = 'Click "Open Report" button to open file report.'


FILE1 = None
FILE2 = None
FILE3 = 'date_hour.xlsx'


def file_1(event):
    """
    :param event:
    :return:FILE1 name
    """
    label1 = tk.Label( WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42',fg='black')
    label1.place(x=700, y=20, width=600, height=85)
    global FILE1
    FILE1 = askopenfilename(
        filetypes=[('FILE1', 'source_file.xlsx'), ('all files', '*.*')])
    if FILE1:
        label1['text'] = str(FILE1) + ' loaded!'
    else:
        label1['fg'] = 'white'
        label1['text'] = 'No file selected! Click "Load" button to select a file.'


def file_2(event):
    """
        :param event:
        :return:FILE2 name
        """
    label2 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='black')
    label2.place(x=800, y=120, width=600, height=85)
    global FILE2
    FILE2 = askopenfilename(
        filetypes=[('FILE2', 'report.xlsx'), ('all files', '*.*')])
    if FILE2:
        label2['text'] = str(FILE2) + ' loaded !'
    else:
        label2['fg'] = 'white'
        label2['text'] = 'No file selected! Click "Report" button to select a file.'


def date_recorder(event):
    """
    :param event:
    :return: record data
    """
    label3 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='black')
    label3.place(x=780, y=220, width=400, height=85)

    # write date and hour:
    WB = openpyxl.load_workbook(FILE3)
    SHEET = WB['Sheet']

    date_now = datetime.now()  # current date and hour
    hour_now = date_now.strftime('%H:%M:%S')  # hour format
    day_now = date_now.strftime('%d''-''%m''-''%y')  # day format
    epoch = time.time()  # epoch time in seconds (from 01.01.1970)

    numero_cell = ['A' + str(i) for i in range(1, 22)]
    hour_cell = ['B' + str(i) for i in range(1, 22)]
    date_cell = ['C' + str(i) for i in range(1, 22)]
    epoch_cell = ['D' + str(i) for i in range(1, 22)]

    count = 1
    delta = 6
    wait_sec = 5

    for i in epoch_cell[1:21]:
        if SHEET[i].value is not None:
            delta=epoch-SHEET[i].value
            count += 1
            if count == 21:
                label3['text'] = 'Full memory! Press delete for reset data!'
        elif delta > wait_sec:
            SHEET[i].value = epoch
            SHEET[date_cell[count]].value = day_now
            SHEET[hour_cell[count]].value = hour_now
            SHEET[numero_cell[count]].value = str(count) + '.'
            label3['text'] = 'Values have been recorded!'
            deg_reg()  # write temperature data
            break
        else:
            label3['text'] = f'Please wait {wait_sec - int(delta)} seconds until the next record! '
            break
    WB.save(FILE3)


def deg_reg():
    """
    :return: temperature data
    """
    global FILE1
    WB = openpyxl.load_workbook(FILE1)
    SHEET = WB['Sheet']
    degree = random.randint(18, 22)
    numero_cell = ['A' + str(i) for i in range(1, 22)]
    degree_cell = ['B' + str(i) for i in range(1, 22)]
    count = 0
    for i in numero_cell[1:21]:
        count += 1
        if SHEET[i].value is None:
            SHEET[i].value = str(count) + '.'
            SHEET[degree_cell[count]].value = int(degree)
            break
    WB.save(FILE1)


def range_letter(start, stop):
    """
    :param start: first letter
    :param stop: last letter
    :return: the character that represents the unicode
    """
    return (chr(n) for n in range(ord(start), ord(stop) + 1))


def execute(event):
    """
    :param event:
    :return: create report
    """
    label4 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='black')
    label4.place(x=750, y=320, width=700, height=85)
    # load and read from file 1
    WB = openpyxl.load_workbook(FILE1)
    SHEET = WB.active
    no_list = []
    t_list = []

    for i in range(1, 22):
        no_cell = SHEET['A' + str(i)].value
        no_list.append(no_cell)
        temp_cell = SHEET['B' + str(i)].value
        t_list.append(temp_cell)

    # load and read from FILE3
    WB = openpyxl.load_workbook(FILE3)
    SHEET = WB.active
    h_list = []
    d_list = []
    for i in range(1, 22):
        hour_cell = SHEET['B' + str(i)].value
        h_list.append(hour_cell)
        date_cell = SHEET['C' + str(i)].value
        d_list.append(date_cell)

    # write report
    line_no = [str(i) + '2' for i in range_letter("A", "Z")]
    line_temp = [str(i) + '3' for i in range_letter("A", "Z")]
    line_hour = [str(i) + '4' for i in range_letter("A", "Z")]
    line_date = [str(i) + '5' for i in range_letter("A", "Z")]
    line_nominal = [str(i) + '8' for i in range_letter("A", "Z")]
    line_l_tol = [str(i) + '9' for i in range_letter("A", "Z")]
    line_u_tol = [str(i) + '10' for i in range_letter("A", "Z")]

    WB = openpyxl.load_workbook(FILE2)
    SHEET = WB.active

    for i in range(21):
        # start_no_list= start index for each list
        start_no_list = no_list[i]
        start_t_list = t_list[i]
        start_h_list = h_list[i]
        start_d_list = d_list[i]

        SHEET[line_no[i]].value = start_no_list
        SHEET[line_temp[i]].value = start_t_list
        SHEET[line_hour[i]].value = start_h_list
        SHEET[line_date[i]].value = start_d_list

    for i in range(1,21):
        if SHEET[line_no[i]].value is not None:
            SHEET[line_nominal[i]].value = 20
            SHEET[line_u_tol[i]].value = 22
            SHEET[line_l_tol[i]].value = 18
        else:
            SHEET[line_nominal[i]].value = None
            SHEET[line_u_tol[i]].value = None
            SHEET[line_l_tol[i]].value = None
    label4['text'] = f'Report created: {FILE2}'
    WB.save(FILE2)


def reset_data(event):
    """
    :param event:
    :return: clear table from all files
    """
    label5 = tk.Label(WINDOW, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='black')
    label5.place(x=780, y=420, width=300, height=85)
    # read and load FILE3
    global FILE3
    WB = openpyxl.load_workbook(FILE3)
    SHEET = WB.active
    # delete columns
    for i in range(4):  # using for loop with range(4) to delete column 1 four times
        SHEET.delete_cols(1)  # when deleted column 1 it's been replaced with column 2 and so on
    # write
    numero = SHEET.cell(row=1, column=1)
    hour = SHEET.cell(row=1, column=2)
    date = SHEET.cell(row=1, column=3)
    time = SHEET.cell(row=1, column=4)

    number_title = 'No.'
    hour_title = 'Hour'
    date_title = 'Date'
    epoch_title = 'Epoch'

    numero.value = number_title
    hour.value = hour_title
    date.value = date_title
    time.value = epoch_title
    WB.save(FILE3)

    global FILE1
    WB = openpyxl.load_workbook(FILE1)
    SHEET = WB.active

    for i in range(2):
        SHEET.delete_cols(1)
        numero = SHEET.cell(row=1, column=1)
        degree = SHEET.cell(row=1, column=2)
        number = 'No.'
        d1 = 'Temperature(°C)'
        numero.value = number
        degree.value = d1
    WB.save(FILE1)
    label5['text'] = 'Data has been deleted!'


def show_report(event):
    if sys.platform == "win32":
        os.startfile(f'{FILE2}')
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, FILE2])


BUTTON_FILE_1.bind("<Button>", file_1)
BUTTON_FILE_2.bind("<Button>", file_2)
BUTTON_OPEN.bind("<Button>", show_report)
BUTTON_EXECUTION.bind("<Button>", execute)
BUTTON_RESET_DATA.bind("<Button>", reset_data)
BUTTON_DATE_RECORDER.bind("<Button>", date_recorder)
WINDOW.mainloop()
