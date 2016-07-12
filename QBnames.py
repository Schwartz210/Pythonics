__author__ = 'aschwartz - Schwartz210@gmail.com'
from mydatabases import clients, jobs, cc_accts, names, general_ledger, addresses
from pyautogui import *
from gui import QuickBooksTransactionInterface
from time import sleep
from spreadsheet import open_file

def get_things_to_enter_into_qb():
    filepath = 'names_to_enter_into_QB.xlsx'
    data = open_file(filepath,0,200)
    return data

def quick_add(button=None):
    if button == None:
        button = 'tab'
    quick = locateOnScreen('img/quick.png')
    quick2 = locateOnScreen('img/quick2.png')
    sleep(1 )
    if quick or quick2:
        press('q')
        sleep(1)
        press(button)
        sleep(1.5)

def incorrect_item_in_date_field():
    not_a_date = locateOnScreen('img/incorrect_date_field.png')
    if not_a_date:
        press('enter')
        press('backspace')
        hotkey('shift','tab')
        sleep(1)

def separate_into_lists():
    enter_into_qb = get_things_to_enter_into_qb()
    vendors = set()
    gl = set()
    jobs = set()
    del enter_into_qb[0]
    for row in enter_into_qb:
        vendors.add(row[0])
        gl.add(row[1])
        jobs.add(row[2])
    return vendors, gl, jobs

def setup_gl():
    setup_prompt = locateOnScreen('img/setup.png'
                                  ''
                                  '')
    if setup_prompt:
        press('s')
        sleep(1)
        press('enter')
        sleep(1)
        press('enter')
        sleep(1)

def main_inputter_subroutine(data=None):
    if data:
        vendors, gl, jobs = data
    else:
        vendors, gl, jobs = separate_into_lists()

    QB_icon_loc, hdrive_loc = QuickBooksTransactionInterface.find_locations_first_screen()
    QuickBooksTransactionInterface.special_click(QB_icon_loc)
    hotkey('alt','t')
    sleep(.5)
    press('down')
    press('enter')
    sleep(1)
    for vendor in vendors:
        print vendor
        #typewrite(vendor)
        #incorrect_item_in_date_field()
        #press('tab')
        #sleep(1)
        #quick_add()
        #hotkey('shift','tab')
        #incorrect_item_in_date_field()

    for i in range(8):
        press('tab')
    for ledger in gl:
        typewrite(ledger)
        press('tab')
        sleep(1)
        setup_gl()
        hotkey('shift','tab')
    for i in range(3):
        press('tab')
    for job in jobs:
        typewrite(job)
        press('down')
        sleep(1)
        quick_add('down')
    for i in range(4):
        press('tab')
    typewrite('operation complete')






