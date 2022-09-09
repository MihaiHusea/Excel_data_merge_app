'''
Proiect excel automation
'''
import datetime
import time
import tkinter as tk
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
from datetime import datetime

file1 = ''
file2 = ''
file3 = 'date_hour.xlsx'


wb = load_workbook(file3)
sheet = wb.active

def execution():
    pass

# grafica

window = tk.Tk()
window.geometry("250x450")
window.title('Excell app')

buton_file_1 = tk.Button(
    text="Browse",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="green",
    fg="black",
)
buton_file_1.pack()

buton_file_2 = tk.Button(
    text="File",
    width=15,
    height=3,
    font=("Arial", 15, "bold"),
    bg="green",
    fg="black",
)
buton_file_2.pack()

buton_execution = tk.Button(
    text="Execution",
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
    bg="blue",
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


def file_1(event):
    global file1
    file1 = askopenfilename(
        filetypes=[('text', '*.xlsx'), ('all files', '*.*')]
    )


def file_2(event):
    global file2
    file2 = askopenfilename(
        filetypes=[('text', '*.xlsx'), ('all files', '*.*')]
    )


def reset_data(event):
    # read and load file
    global file3
    print('Read and load ' + file3 + '...')
    wb = load_workbook(file3)
    sheet = wb.active
    # delete columns
    for i in range(4):  # in loc sa punem de 4 ori sheet.delete_cols(1)
        # facem un for cu range de 4 , stergem coloana 1 de 4 ori
        sheet.delete_cols(1)  # de fiecare data cand stergem coloana 1 este inlocuita cu urmatoarea,
        # de aceea trebuie sa stergem de 4 ori
    print('Delete ' + file3 + '...')
    # write
    numero=sheet.cell(row=1, column=1)
    hour = sheet.cell(row=1, column=2)
    date = sheet.cell(row=1, column=3)
    time = sheet.cell(row=1, column=4)

    no='No.'
    h1 = 'Hour'
    d1 = 'Date'
    t1 = 'Epoch'
    # t2=0 #folosit pentru range
    numero.value=no
    hour.value = h1
    date.value = d1
    time.value = t1
    print('Write ' + file3 + '...')
    wb.save(file3)
    print('Save ' + file3 + '...')


def date_recorder(event):
    global file3

    # load and write:
    wb = load_workbook(file3)
    sheet = wb['Sheet']

    date_now = datetime.now()                   # data si ora actuala
    hour_now = date_now.strftime('%H:%M:%S')    # ora actuala
    day_now = date_now.strftime('%d''-''%m''-''%y')    # ziua actuala
    epoch = time.time()                         # timpul secunde


    numero_cell=['A' + str(i) for i in range(1, 22)]
    hour_cell = ['B' + str(i) for i in range(1, 22)]
    date_cell = ['C' + str(i) for i in range(1, 22)]
    epoch_cell = ['D' + str(i) for i in range(1, 22)]

    x = 1
    delta = 6
    wait_sec = 5

    for i in epoch_cell[1:21]:
        if sheet[i].value is not None:
            t1 = float(sheet[i].value)
            delta = float(epoch)-t1
            x += 1
            if x == 21:
                print('toate celulele au fost completate! folositi reset pentru rescriere')
            continue

        elif delta > wait_sec:
            sheet[i].value = epoch
            sheet[date_cell[x]].value = day_now
            sheet[hour_cell[x]].value = hour_now
            sheet[numero_cell[x]].value=str(x)+'.'
            print('values has been recorded...')
            break
        else:
            print(f'Asteptati {wait_sec - int(delta)} secunde pana la inregistrarea urmatoare ')
            break
    wb.save(file3)


buton_file_1.bind("<Button-1>", file_1)
buton_file_2.bind("<Button-1>", file_2)
buton_execution.bind("<Button-1>", execution)
buton_reset_data.bind("<Button-1>", reset_data)
buton_date_recorder.bind("<Button-1>", date_recorder)

window.mainloop()

# todo de verificat cu pylint
# todo de redenumit var sau alti termeni in engleza
# todo de pus descriere pentru fiecare functie cu ''' '''
# todo wait time de intrudus manual (buton sau alta varianta)
# todo de verificat implemantare elif in loc de if la data
# todo de adaptat codul la oop
# todo de verificat daca aplicatia poate fi setata in fereastra principala
# todo de facut functie pentru parcurgerea second cell si calcul delta
# todo de parcurs codul in vederea aflarii posibilelor intrebari
#todo de scris epoch excel cu alta culoare
#todo de implementat  apelarea a 2 functii pe un buton