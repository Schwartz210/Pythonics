__author__ = 'aschwartz - Schwartz210@gmail.com'
from pyautogui import *
from time import sleep
from class_mod2 import ContainerTransaction, CheckRequest, Invoice, LineItem, LineItemInv, LineItemChk
from spreadsheet import check_request_pull, invoice_pull, open_file, screen_out_empty_rows, screen_out_col, format_date
from error_checking import ErrorChecker
from folder_contents import financial

class QuickBooksTransactionInterface(object):
    def __init__(self, run):
        self.run = run
        self.trans_type = 'Parent'
        self.source = []       #this attribute is overwritten by child class
        self.trans_class = ContainerTransaction  #'Invoice', 'CheckRequest'
        self.line_class = LineItem   #'LineItem', 'LineItemInv'

    def second(self):
        print('OP: function second started')
        self.transaction_list = self.factory()
        self.general_ledger_reorder(self.transaction_list)
        e = ErrorChecker(self.transaction_list, self.trans_type)
        e.print_transactions()
        if self.run: self.qb_interface()

    @staticmethod
    def special_click(thing_to_click):
        click(thing_to_click[0],thing_to_click[1],button='left')

    @staticmethod
    def multi_tab(qty):
        for i in range(qty):
            press('tab')

    @staticmethod
    def find_locations_first_screen():
        QB_icon_loc = locateOnScreen('img/qb.png')
        hdrive_loc = locateOnScreen('img/hdrive.png')
        return QB_icon_loc, hdrive_loc

    def is_COGS(self, general_ledger):
        if general_ledger[:7] == 'Meeting':
            return True
        else:
            return False

    def general_ledger_reorder(self, transactions):
        for transaction in transactions:
            new_line_items = []
            for line_item in transaction.line_items:
                if self.is_COGS(line_item.gl_out):
                    new_line_items.append(line_item)
                else:
                    new_line_items.insert(0, line_item)
                transaction.line_items = new_line_items

    def factory(self):
        '''
        This function creates Invoices, CheckRequests, LineItems obj's and packages them all up.
        '''
        line_num = 1
        iterator = 1
        transactions = []
        for packet in self.source:
            print('OP: factory method in prog. Packet %s of %s' % (iterator, len(self.source)))
            iterator += 1
            if packet[0] == 'WHOLE':
                line_item = self.line_class(packet)
                line_item.line_num = line_num
                trasaction = self.trans_class(packet)
                trasaction.append(line_item)
                transactions.append(trasaction)
            elif packet[0] == 'TOP':
                trasaction = self.trans_class(packet)
                trasaction.line_num = line_num
                transactions.append(trasaction)
            elif packet[0] == 'SPLT':
                line_item = self.line_class(packet)
                line_item.line_num = line_num
                transactions[-1].append(line_item)
            line_num += 1
        print('OP: factory method complete')
        return transactions

    def qb_interface(self):
        '''
        Overwritten by child classes
        '''
        pass

    def transaction_level_subroutine(self, transaction_list):
        for transaction in self.transaction_list:
            transaction.QB()
            sleep(3)

class InvoiceInterface(QuickBooksTransactionInterface):
    def __init__(self, run):
        QuickBooksTransactionInterface.__init__(self, run)
        self.trans_type = 'invoice_list'
        self.source = invoice_pull()
        self.trans_class = Invoice
        self.line_class = LineItemInv
        self.second()

    def qb_interface(self):
        QB_icon_loc, hdrive_loc = self.find_locations_first_screen()
        self.special_click(QB_icon_loc)
        hotkey('alt','t')
        sleep(.5)
        press('down')
        press('enter')
        sleep(1)
        self.transaction_level_subroutine(self.transaction_list)


class CheckInterface(QuickBooksTransactionInterface):
    def __init__(self, run):
        QuickBooksTransactionInterface.__init__(self, run)
        self.trans_type = 'check'
        self.source = check_request_pull()
        self.trans_class = CheckRequest
        self.line_class = LineItemChk
        self.second()

    def qb_interface(self):
        QB_icon_loc, hdrive_loc = self.find_locations_first_screen()
        self.special_click(QB_icon_loc)
        hotkey('alt','b')
        sleep(.5)
        press('enter')
        sleep(2)
        hotkey('shift','tab')
        typewrite('t')
        press('tab')
        self.transaction_level_subroutine(self.transaction_list)


class AllocationInterface(InvoiceInterface):
    def __init__(self, run, spreadsheet_tab):
        self.run = run
        self.spreadsheet_tab = spreadsheet_tab
        self.trans_type = 'allocation_spread'
        self.source = self.allocation_pull()
        self.trans_class = Invoice
        self.line_class = LineItemInv
        self.spreadsheet_tab = spreadsheet_tab
        self.second()


    def allo_spreadsheet_screen_out_content(self, list):
        result = screen_out_empty_rows(list)
        result = screen_out_col(result, 4, 9)
        return result

    def packet_converter(self, data_list):
        '''
        Receives as strangely ordered data set from the spreadsheet (has multiple rows of headers before the line_items)-
        and converts it to the same field format as checks or invoices.
        '''
        output_list = []
        splt_data = list(data_list)
        for i in range(5):
            del splt_data[0]

        vendor = data_list[0][1]
        gl_in_general = data_list[0][4].strip(':')
        inv_date = data_list[1][1]
        ref_num = data_list[2][1]
        total_amount = data_list[3][1]
        first_row = ['TOP','','-','Total',vendor,inv_date,ref_num,total_amount,ref_num + '.pdf']
        output_list.append(first_row)
        for splt in splt_data:
            if splt[2] == '':
                gl = gl_in_general
            else:
                gl = splt[2]
            comment = splt[3] + " - " + format_date(splt[4])
            row = ['SPLT',splt[1],gl,comment,vendor,inv_date,ref_num,splt[0],ref_num + '.pdf']
            output_list.append(row)
        return output_list

    def allocation_pull(self):
        invoice_allocation = financial + r'\Avi\1. Invoice worksheet - Allocation.xlsm'
        allocation_spread = open_file(invoice_allocation,self.spreadsheet_tab,100)
        allocation_spread = self.allo_spreadsheet_screen_out_content(allocation_spread)
        allocation_spread = self.packet_converter(allocation_spread)
        return allocation_spread

