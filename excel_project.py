#!/usr/bin/env python3

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
window = Tk()
window.title('AppX v1.0')
window.geometry("1400x650")
window.configure(background="#1D6F42")
background_text = tk.Label(window,
                           text='KEEP\nCALM\nIT' + "'" + 'S JUST AN\nEXCEL\nFILE',
                           bg='#1D6F42',
                           font=("Arial", 35, "bold"),
                           fg="white")
background_text.place(x=20, y=165)

button_file_1 = tk.Button(
    text="Load",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg='#abf77e',
    fg="black",
    highlightthickness=0
)

button_file_1.place(x=450, y=20)
label1 = Label(window, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label1.place(x=700, y=20, width=600, height=85)
label1['text'] = 'No file selected! Click "Load" button to select a file.'

button_file_2 = tk.Button(
    text="Report",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#abf77e",
    fg="black",
    highlightthickness=0
)
button_file_2.place(x=550, y=120)
label2 = tk.Label(window, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label2.place(x=800, y=120, width=600, height=85)
label2['text'] = 'No file selected! Click "Report" button to select a file.'

button_date_recorder = tk.Button(
    text="Rec",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#abf77e",
    fg="black",
    highlightthickness=0
)
button_date_recorder.place(x=450, y=220)
label3 = tk.Label(window, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label3.place(x=650, y=220, width=600, height=85)
label3['text'] = 'Click "Rec" button to record data.'

button_execution = tk.Button(
    text="Execute",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#3EB489",
    fg="black",
    highlightthickness=0
)
button_execution.place(x=550, y=320)
label4 = tk.Label(window, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label4.place(x=750, y=320, width=600, height=85)
label4['text'] = 'Click "Execute" button to create data report.'

button_reset_data = tk.Button(
    text="Delete",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="#E00201",
    fg="black",
    highlightthickness=0
)
button_reset_data.place(x=450, y=420)
label5 = tk.Label(window, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label5.place(x=650, y=420, width=600, height=85)
label5['text'] = 'Click "Delete" button to reset data.'

button_open = tk.Button(
    text="Open Report",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="yellow",
    fg="black",
    highlightthickness=0
)
button_open.place(x=550, y=520)
label6 = tk.Label(window, font=('Arial', 12, 'bold'), bg='#1D6F42', fg='white')
label6.place(x=750, y=520, width=600, height=85)
label6['text'] = 'Click "Open Report" button to open file report.'

file1 = None
file2 = None
file3 = 'date_hour.xlsx'


def select_source_file(event):
    """
    :param event:
    :return:FILE1 name
    """
    global file1
    label1['fg'] = '#1D6F42'
    file1 = askopenfilename(
        filetypes=[('FILE1', 'source_file.xlsx'), ('all files', '*.*')])
    label1['fg'] = 'black'
    if file1:
        label1['text'] = str(file1) + ' loaded!'
    else:
        label1['fg'] = 'white'
        label1['text'] = 'No file selected! Click "Load" button to select a file.'


def select_report_file(event):
    """
        :param event:
        :return:FILE2 name
        """
    global file2
    label2['fg'] = '#1D6F42'
    file2 = askopenfilename(
        filetypes=[('FILE2', 'report.xlsx'), ('all files', '*.*')])
    label2['fg'] = 'black'
    if file2:
        label2['text'] = str(file2) + ' loaded !'
    else:
        label2['fg'] = 'white'
        label2['text'] = 'No file selected! Click "Report" button to select a file.'


def record_data(event):
    """
    :param event:
    :return: record data
    """
    label3['fg'] = 'black'
    # write date and hour:
    wb = openpyxl.load_workbook(file3)
    sheet = wb['Sheet']

    date_now = datetime.now()  # current date and hour
    hour_now = date_now.strftime('%H:%M:%S')  # hour format
    day_now = date_now.strftime('%d''-''%m''-''%y')  # day format
    epoch = time.time()  # epoch time in seconds (from 01.01.1970)

    number_cell = ['A' + str(i) for i in range(1, 22)]
    hour_cell = ['B' + str(i) for i in range(1, 22)]
    date_cell = ['C' + str(i) for i in range(1, 22)]
    epoch_cell = ['D' + str(i) for i in range(1, 22)]

    count = 1
    delta = 6
    wait_sec = 5

    for i in epoch_cell[1:21]:
        if sheet[i].value is not None:
            delta = epoch - sheet[i].value
            count += 1
            if count == 21:
                label3['text'] = 'Full memory! Press delete for reset data!'
        elif delta > wait_sec:
            sheet[i].value = epoch
            sheet[date_cell[count]].value = day_now
            sheet[hour_cell[count]].value = hour_now
            sheet[number_cell[count]].value = str(count) + '.'
            label3['text'] = 'Values have been recorded!'
            record_temperature()  # write temperature data
            break
        else:
            label3['text'] = f'Please wait {wait_sec - int(delta)} seconds until the next record! '
            break
    wb.save(file3)


def record_temperature():
    """
    :return: temperature data
    """
    global file1
    wb = openpyxl.load_workbook(file1)
    sheet = wb['Sheet']
    degree = random.randint(18, 22)
    number_cell = ['A' + str(i) for i in range(1, 22)]
    degree_cell = ['B' + str(i) for i in range(1, 22)]
    count = 0
    for i in number_cell[1:21]:
        count += 1
        if sheet[i].value is None:
            sheet[i].value = str(count) + '.'
            sheet[degree_cell[count]].value = int(degree)
            break
    wb.save(file1)


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
    label4['fg'] = 'black'
    # load and read from file 1
    wb = openpyxl.load_workbook(file1)
    sheet = wb.active
    no_list = []
    t_list = []

    for i in range(1, 22):
        no_cell = sheet['A' + str(i)].value
        no_list.append(no_cell)
        temp_cell = sheet['B' + str(i)].value
        t_list.append(temp_cell)

    # load and read from FILE3
    wb = openpyxl.load_workbook(file3)
    sheet = wb.active
    h_list = []
    d_list = []
    for i in range(1, 22):
        hour_cell = sheet['B' + str(i)].value
        h_list.append(hour_cell)
        date_cell = sheet['C' + str(i)].value
        d_list.append(date_cell)

    # write report
    line_no = [str(i) + '2' for i in range_letter("A", "Z")]
    line_temp = [str(i) + '3' for i in range_letter("A", "Z")]
    line_hour = [str(i) + '4' for i in range_letter("A", "Z")]
    line_date = [str(i) + '5' for i in range_letter("A", "Z")]
    line_nominal = [str(i) + '8' for i in range_letter("A", "Z")]
    line_l_tol = [str(i) + '9' for i in range_letter("A", "Z")]
    line_u_tol = [str(i) + '10' for i in range_letter("A", "Z")]

    wb = openpyxl.load_workbook(file2)
    sheet = wb.active

    for i in range(21):
        # start_no_list= start index for each list
        start_no_list = no_list[i]
        start_t_list = t_list[i]
        start_h_list = h_list[i]
        start_d_list = d_list[i]

        sheet[line_no[i]].value = start_no_list
        sheet[line_temp[i]].value = start_t_list
        sheet[line_hour[i]].value = start_h_list
        sheet[line_date[i]].value = start_d_list

    for i in range(1, 21):
        if sheet[line_no[i]].value is not None:
            sheet[line_nominal[i]].value = 20
            sheet[line_u_tol[i]].value = 22
            sheet[line_l_tol[i]].value = 18
        else:
            sheet[line_nominal[i]].value = None
            sheet[line_u_tol[i]].value = None
            sheet[line_l_tol[i]].value = None
    label4['text'] = f'Report created: {file2}'
    wb.save(file2)


def reset_data(event):
    """
    :param event:
    :return: clear table from all files
    """
    label5['fg'] = 'black'
    # read and load FILE3
    global file3
    wb = openpyxl.load_workbook(file3)
    sheet = wb.active
    # delete columns
    for i in range(4):  # using for loop with range(4) to delete column 1 four times
        sheet.delete_cols(1)  # when deleted column 1 it's been replaced with column 2 and so on
    # write
    number = sheet.cell(row=1, column=1)
    hour = sheet.cell(row=1, column=2)
    date = sheet.cell(row=1, column=3)
    time = sheet.cell(row=1, column=4)

    number_title = 'No.'
    hour_title = 'Hour'
    date_title = 'Date'
    epoch_title = 'Epoch'

    number.value = number_title
    hour.value = hour_title
    date.value = date_title
    time.value = epoch_title
    wb.save(file3)

    global file1
    wb = openpyxl.load_workbook(file1)
    sheet = wb.active

    for i in range(2):
        sheet.delete_cols(1)
        number = sheet.cell(row=1, column=1)
        degree = sheet.cell(row=1, column=2)
        number = 'No.'
        d1 = 'Temperature(°C)'
        number.value = number
        degree.value = d1
    wb.save(file1)
    label5['text'] = 'Data has been deleted!'


def show_report(event):
    if sys.platform == "win32":
        os.startfile(f'{file2}')
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, file2])


button_file_1.bind("<Button>", select_source_file)
button_file_2.bind("<Button>", select_report_file)
button_open.bind("<Button>", show_report)
button_execution.bind("<Button>", execute)
button_reset_data.bind("<Button>", reset_data)
button_date_recorder.bind("<Button>", record_data)
window.mainloop()
