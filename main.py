# This is a sample Python script.
#
# # Press Shift+F10 to execute it or replace it with your code.
# # Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
#
#
# def print_hi(name):
#     # Use a breakpoint in the code line below to debug your script.
#     print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
#
#
# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')
#
# # See PyCharm help at https://www.jetbrains.com/help/pycharm/


# hour_cell = ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11',
#                  'B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21']
# date_cell = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12',
#              'C13', 'C14', 'C15', 'C16', 'C17', 'C18', 'C19', 'C20', 'C21']
# second_cell = ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10', 'D11',
#                'D12', 'D13', 'D14', 'D15', 'D16', 'D17', 'D18', 'D19', 'D20', 'D21']

# oran = date_now.strftime('%H'':''%M'':''%S')


#


#
# # difference between dates in timedelta
# delta = d2 - d1
# print(f'Difference is {delta.seconds} seconds ')
#
# date_now = datetime.datetime.now()
# print(date_now)

# from datetime import datetime
# dt = datetime.today()  # Get timezone naive now
# seconds = dt.timestamp()
# print(dt)
# print(seconds)
# import time
#
# a=time.time()
# print(a)

delta = '00:00:06'
wait_sec = '00:00:05'

d1=delta.split(':')
print(d1)

w1=wait_sec.split(':')
print(w1)

print(int(d1[2])>int(w1[2]))