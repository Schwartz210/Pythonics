__author__ = 'aschwartz - Schwartz210@gmail.com'
from gui import InvoiceInterface, CheckInterface, AllocationInterface
from QBnames import main_inputter_subroutine
from folder_contents import main_file_handler, rename_move_files, show_folder_contents, open_folder, open_file2, open_amex
from startup import Startup
from logger import Logger
from iif_class import AmexFactory




def main():
    #AllocationInterface(True, spreadsheet_tab='5-476-52387')
    InvoiceInterface(True)
    #CheckInterface(True)
    #main_inputter_subroutine()
    main_file_handler('h','payables')
    #rename_move_files('h','payables', 'Inv-07.12.16 - ')
    #show_database('GL')
    #show_folder_contents('h')
    #open_folder('root')
    #open_folder('payables')
    #open_file2('invoices')
    #open_amex()
    #Startup()
    #AmexFactory()
    print '\n\n\n*** PROCESS COMPLETE ***'



















main()



