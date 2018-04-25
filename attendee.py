'''
This module reads an attendee report and uses pyautogui to directly input data into QuickBooks
'''
from spreadsheet import open_file, screen_out_empty_rows, keep_these_columns
from pyautogui import *
from time import sleep

def get_data(path):
    data = open_file(path, 1, 100)
    data = keep_these_columns(data,8,9,10,11,12)
    data = screen_out_empty_rows(data)
    del data[0]
    return data

def clean(record):
    if 'Deposit' in record[2]:
        record[4] = '^ Deposit'
    if 'Room Rental' in record[0] or 'AudioVisual' in record[0]:
        record[4] = ''

def slow_typing(word):
    for letter in word[:5]:
        typewrite(letter)
        sleep(.2)
    typewrite(word[5:])

def modified_typewrite(string):
    if string == '':
        for _ in range(10):
            press('backspace')
    else:
        typewrite(string)

def input_record(record):
    slow_typing(record[0])
    press('tab')
    sleep(1)
    typewrite(str(record[1]))
    press('tab')
    sleep(1)
    typewrite(record[2])
    press('tab')
    sleep(1)
    typewrite(record[3])
    press('tab')
    sleep(1)
    quick = locateOnScreen('img/quick.png')
    quick2 = locateOnScreen('img/quick2.png')
    if quick or quick2:
        press('q')
    sleep(1.5)
    press('tab')
    modified_typewrite(record[4])
    sleep(1)
    press('tab')
    sleep(1.5)
    quick = locateOnScreen('img/quick.png')
    quick2 = locateOnScreen('img/quick2.png')
    if quick or quick2:
        press('q')
    sleep(1.5)

def wait(seconds):
    counter = seconds
    for _ in range(seconds):
        print(counter)
        sleep(1)
        counter -= 1

def save_exit():
    hotkey('alt', 'a')
    sleep(.25)
    press('enter')
    print('Done')

def execute(path):
    wait(5)
    data = get_data(path)
    for record in data:
        clean(record)
        input_record(record)
    save_exit()



