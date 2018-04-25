__author__ = 'aschwartz - Schwartz210@gmail.com'
from spreadsheet import invoice_pull, check_request_pull
from sql_time import pull_data, exists
from super_transaction import SuperTransaction
from pyautogui import *
from time import sleep
from datetime import datetime
from folder_contents import folder_dict
from subjobs import subjob_builder


class TransactionSegment(SuperTransaction):
    '''
    This class is the parent class for all transaction classes used for 1) invoice runs, and 2) check runs.
    Child classes include 'ContainerTransaction'(Invoices, CheckRequests), which hold their respective line item classes.
    This class pulls data from the spreadsheets 'Mass check request gui.xlsx' and 'Mass invoice gui.xlsx'. This data is-
    mapped by row and column to class attributes.
    In accounting terms, this class is used to represent both the debit/credit side, and all associated details.
    '''
    def __init__(self, packet):
        SuperTransaction.__init__(self)
        self.source = ''
        self.type_info = packet[0]
        self.job_in = self.str_format(packet[1])
        self.gl_in = packet[2]
        self.comments = packet[3]
        self.vendor_name = packet[4]
        self.date = self.format_date(packet[5])
        self.ref_num = self.clean_ref_num(packet[6])
        self.amount = packet[7]
        self.doc_name = str(packet[8])
        self.line_num = 0
        self.directory = self.create_directory()
        self.link = self.create_link(self.directory, self.doc_name)

    def find_location(self):
        sql_request = 'SELECT location FROM jobs WHERE number="%s"' % (self.job_in)
        data = pull_data(sql_request)
        try:
            return data[0][0]
        except:
            subjob_builder([self.job_in])
            return None

    def find_program_date(self):
        sql_request = 'SELECT date FROM jobs WHERE number="%s"' % (self.job_in)
        if exists(sql_request):
            data = pull_data(sql_request)
            return data[0][0]
        else:
            return None

    def bold(self, string):
        '''
        Makes string bold
        '''
        return '\033[1m' + string + '\033[0m'

    def clean_ref_num(self, raw_ref_num):
        '''
        Cleans up formatting
        '''
        str_ref_num = str(raw_ref_num)
        if str_ref_num and str_ref_num[0] == '?':
            # '?' is used to prevent some invoice numbers as being read as dates instead of normal string
            str_ref_num = str_ref_num[1:]

        if str_ref_num[-2:] == '.0':
            #Sometimes invoice numbers are wrongfully converted in floats. This converts it back to its intended state.
            formatted_ref_num = str_ref_num[:len(str_ref_num) - 2]
            return formatted_ref_num
        else:
            return str_ref_num

    def create_directory(self):
        '''
        Just a little bit of string formatting.
        '''
        snip = folder_dict['payables']
        snip = snip.strip('S:')
        start = r'\\crgnas01\shared'
        return start + snip

class LineItem(TransactionSegment):
    '''
    This class is for line items within a ContainerTransaction.
    In accounting terms, it represents the 'splits' or line items (the bottom portion) of the transaction)
    '''
    def __init__(self, packet):
        TransactionSegment.__init__(self,packet)
        if not self.gl_in:
            self.gl_in = self.override_gl_in()
        self.vendor_obj = self.create_vendor_obj()
        self.full_job = self.find_full_job_name(self.job_in)
        self.short_job = self.find_short_job(self.job_in)
        self.job_out = self.create_job_out(self.full_job, self.gl_in)
        self.client = self.find_client(self.full_job)
        self.gl_out = self.create_gl_out(self.gl_in, self.client)
        self.program_date = self.find_program_date()
        self.memo = self.create_memo()
        self.tran_class = self.create_trans_class()
        self.tabulate_box = [self.gl_out,str(self.amount),self.memo,self.job_out]

    def create_memo(self):
        '''
        This method returns a string that becomes the memo line of a line item.
        '''
        try:
            loc = self.find_location()
        except:
            return None
        if type(self.vendor_obj) == str:
            print(self.vendor_name)
            raise Exception('Vendor object not formed. Update vendor address report')
        try:
            return self.vendor_obj.first_name + ' ' + self.vendor_obj.last_name + ' - ' + self.comments + ' - ' + str(loc) + ' - ' + str(self.program_date)
        except:
            return None

    def create_trans_class(self):
        '''
        'Class' is a field in QB, which we use to record the person a travel charge is associated with.
        The format is 'first_name, last_name'
        '''
        if not self.vendor_obj:
            pass
        elif self.vendor_obj.last_name and self.vendor_obj.first_name:
            return self.vendor_obj.last_name + ', ' + self.vendor_obj.first_name
        else:
            return None

    def QB(self):
        '''
        This method is the GUI interface with QB. When invoked this method enters data into QB.
        pyautogui and time are the libraries used for this.
        '''
        typewrite(self.gl_out)
        press('tab')
        sleep(1)
        typewrite(str(self.amount))
        press('tab')
        sleep(1)
        typewrite(self.memo)
        press('tab')
        sleep(1)
        if self.job_out:
            typewrite(self.job_out)
            press('tab')
        sleep(1)
        press('q')
        sleep(.5)
        press('backspace')
        sleep(1.5)
        press('tab')
        if self.tran_class:
            typewrite(self.tran_class)
        sleep(1)
        press('tab')
        sleep(1.5)

    def __str__(self):
        print '______________line item______________'
        print self.bold('General ledger: ') + ' %s' % (self.gl_out)
        print self.bold('Amount        : ') + ' %s' % (self.amount)
        print self.bold('Memo          : ') + ' %s' % (self.memo)
        print self.bold('Job           : ') + ' %s' % (self.job_out)
        print self.bold('Class         : ') + ' %s' % (self.tran_class)

    def override_gl_in(self):
        '''
        This method converts string comments (spreadsheet, Column D) into gl_in's.
        EXAMPLE: 'Mileage' --> 'Ground'
        With most transaction gl_in provides enough detail, however for accounting purposes, more detail is required for-
        the line items of check requests.
        '''
        comments_dict = {'Mileage' : 'Ground',
                        'Parking' : 'Ground',
                        'Tolls' : 'Ground',
                        'Car Service' : 'Ground',
                        'Gratuities' : 'Per Diem',
                        'Meals' : 'Per Diem',
                        'Airfare' : 'Airfare',
                        'Fee For Service' : 'Fee For Service',
                        'Lodging' :'Lodging',
                        'Internet' : 'Per Diem',
                        'Fees' : 'Fees',
                        'Baggage Fee':'Per Diem',
                        'Train' : 'Ground',
                        'Fuel' : 'Ground',
                        'Ground' : 'Ground',
                        'Production' : 'Production',
                        'Per Diem' : 'Per Diem'}
        return comments_dict[self.comments]

class LineItemChk(LineItem):
    '''
    This class is a type of line item. It only contains a method override for a string format.
    '''
    def __init__(self, packet):
        LineItem.__init__(self, packet)
        self.job_out = self.create_job_out(self.full_job, self.gl_in)
        self.gl_out = self.create_gl_out(self.gl_in, self.client)
        self.tabulate_box = [self.gl_out,str(self.amount),self.memo,self.job_out,self.tran_class]




class LineItemInv(LineItem):
    '''
    This class is a type of line item. It only contains a method override for a string format.
    '''
    def __init__(self, packet):
        LineItem.__init__(self, packet)

    def create_memo(self):
        '''
        String formatting for memo field.
        '''
        try:
            if not self.vendor_obj:
                pass
            elif self.comments:
                return self.vendor_obj.payable + ' - ' + self.gl_in + ' - ' + self.comments
            else:
                return str(self.vendor_obj.payable) + ' - ' + str(self.gl_in)
        except:
            raise Exception('Try running a Vendor Address report from QuickBooks.')

class ContainerTransaction(TransactionSegment):
    '''
    This class represents two types of transactions; 1) Invoices and 2) CheckRequest.
    This is called 'ContainerTransaction' because it contains line items.
    '''
    def __init__(self, packet):
        TransactionSegment.__init__(self, packet)
        self.vendor_obj = self.create_vendor_obj()
        self.line_items = []
        self.iterator = 0

    def next(self):
        if self.iterator == len(self.line_items):
            raise StopIteration
        current_item = self.line_items[self.iterator]
        self.iterator += 1
        return current_item

    def __iter__(self):
        return self

    def len(self):
        return len(self.line_items)

    def slow_typing(self, string):
        '''
        This method is designed to overcome an issue with interfacing with QuickBooks vendor field. The issue was that-
        'typewrite' would type the string too fast, and the field in QuickBooks has an auto-populate feature. This-
        would cause the wrong vendor to be entered. By slowing down the entry of the first five characters this issue-
        was resolved.
        '''
        for letter in string[:5]:
            typewrite(letter)
            sleep(.2)
        typewrite(string[5:])

    def append(self, line_item):
        self.line_items.append(line_item)


class Invoice(ContainerTransaction):
    '''
    This class represents Invoice transaction type. It is a container class that holds line items.
    It has a method that controls the keyboard for data entry in QuickBooks.
    '''
    def __init__(self, packet):
        ContainerTransaction.__init__(self,packet)
        self.qb_date = self.find_quickbooks_date(self.date)
        self.tabulate_box = [self.vendor_name, self.qb_date, self.ref_num,str(self.amount)]

    def QB(self):
        '''
        This method interfaces with QuickBooks by taking over keyboard control via pyautogui module.
        It is ran only for invoices and not check requests, because the QuickBooks screens are different enough.
        It is performs data entry.
        '''
        self.slow_typing(self.vendor_name)
        sleep(1.5)
        press('tab')
        sleep(1)
        typewrite(self.qb_date)
        press('tab')
        sleep(1)
        typewrite(self.ref_num)
        press('tab')
        sleep(1)
        typewrite(str(self.amount))
        press('tab')
        press('tab')
        press('tab')
        press('tab')
        press('tab')
        for line_item in self.line_items:
            line_item.QB()
        press('tab')
        press('tab')
        typewrite(self.link)
        sleep(3)
        hotkey('alt','s')

    def __str__(self):
        print '______________TRANSACTION TOP______________'
        print self.bold('Vendor        : ') + ' %s' % (self.vendor_name)
        print self.bold('Date          : ') + ' %s' % (self.qb_date)
        print self.bold('Invoice Num   : ') + ' %s' % (self.ref_num)
        print self.bold('Amount        : ') + ' %s' % (self.amount)
        print self.bold('Support doc   : ') + ' %s' % (self.link)


class CheckRequest(ContainerTransaction):
    '''
    This class represents CheckRequest transaction type. It is a container class that holds line items.
    It has a method that controls the keyboard for data entry in QuickBooks.
    '''
    def __init__(self, packet):
        ContainerTransaction.__init__(self,packet)
        self.memo = self.create_memo()
        self.qb_date = self.create_qb_date()
        self.tabulate_box = [self.vendor_name, self.qb_date, self.memo,str(self.amount)]

    def get_checktype(self):
        if self.comments == 'Fee For Service':
            checktype = 'Fee For Service'
        else:
            checktype = 'Expenses'
        return checktype

    def create_memo(self):
        '''
        String formatting for the 'memo' field.
        '''
        checktype = self.get_checktype()
        if not self.vendor_obj.last_name:
            print(self.vendor_obj.vendor)
            raise Exception('Try running a Vendor Address report from QuickBooks.')
        name = self.vendor_obj.last_name
        doctor_credentials = ['MD', 'PhD']
        if self.vendor_obj.salutation in doctor_credentials:
            name = 'Dr. ' + name
        location = self.find_location()
        program_date = self.find_program_date()
        memo = name + ' - ' + checktype + ' - ' + str(location) + ' - ' + str(program_date)
        return memo

    def create_qb_date(self):
        '''
        This method returns today's date as the check's date. All CheckRequests must be entered in QuickBooks with-
        today's date. This overrides the null from the spreadsheet.
        '''
        now = datetime.now()
        today = now.strftime('%m/%d/%Y')
        return today

    def QB(self):
        '''
        This method controls the keyboard and conducts data entry in QuickBooks. The functions 'press', 'typewrite', -
        and 'hotkey' come from the 'pyautogui' package.
        Notice the for-loop used to loop through line items.
        '''
        self.slow_typing(self.vendor_name)
        sleep(1)
        press('tab')
        sleep(1)
        typewrite(str(self.amount))
        press('tab')
        press('tab')
        press('tab')
        sleep(1)
        typewrite(self.memo)
        sleep(1)
        press('tab')
        for line_item in self.line_items:
            line_item.QB()
        press('tab')
        press('tab')
        count = 17 - len(self.line_items)
        if count < 1:
            count = 1
        for _ in range(count):
            press('down')
            sleep(.25)
        typewrite(self.link)
        sleep(3)
        hotkey('alt','s')

    def __str__(self):
        print '______________TRANSACTION TOP______________'
        print self.bold('Vender        : ') + ' %s' % (self.vendor_name)
        print self.bold('Amount        : ') + ' %s' % (self.amount)
        print self.bold('Memo          : ') + ' %s' % (self.memo)
        print self.bold('Support doc   : ') + ' %s' % (self.link)



def factory(desired_quantity, some_class):
    '''
    Make lots of some_class
    '''
    inventory = []
    for _ in range(desired_quantity):
        thing = some_class()
        inventory.append(thing)
    return inventory