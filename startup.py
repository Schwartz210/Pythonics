from os import startfile
from time import sleep
from pyautogui import typewrite, press, hotkey, FAILSAFE
from folder_contents import open_folder, open_file2
from logger import Logger
from emails import emails


FAILSAFE = False
class Startup(object):
    '''
    When class is initiated the desktop is prepared for the user. QuickBooks, Outlook, and other stuff runs.
    '''
    def __init__(self, password=None):
        '''
        Container function. Holders all the other startup functions.
        '''
        self.set_password(password)
        self.quickbooks()
        self.gmail()
        self.outlook()
        self.files_and_folders()
        logger_message = 'startup initiated'
        Logger('test',logger_message)

    def set_password(self, password):
        if password:
            self.password = password
        else:
            self.password = raw_input('QuickBooks Password: ')

    @staticmethod
    def excel():
        '''
        Starts Microsoft Excel.
        '''
        startfile('C:\Program Files (x86)\Microsoft Office\Office12\EXCEL.EXE')

    def quickbooks(self):
        '''
        Starts QuickBooks. Logs in.
        '''
        startfile('"C:\Program Files (x86)\Intuit\QuickBooks Enterprise Solutions 16.0\QBW32Enterprise.exe"')
        sleep(30)
        typewrite(self.password)
        press('enter')
        sleep(5)

    def outlook(self):
        '''
        Starts Outlook.
        '''
        startfile('C:\Program Files (x86)\Microsoft Office\Office12\OUTLOOK.EXE')
        sleep(5)

    def files_and_folders(self):
        '''
        Opens specified files and folders.
        '''
        open_folder('h')
        open_file2('jobs')

    def gmail(self):
        '''
        Starts Firefox. Navigates to Gmail.
        '''
        startfile(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
        sleep(5)
        hotkey('ctrl','l')
        sleep(2)
        typewrite('www.gmail.com')
        press('enter')
        sleep(2)