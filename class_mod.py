__author__ = 'aschwartz - Schwartz210@gmail.com'
from spreadsheet import invoice_list, checkrequest_list, gl_list, name_class_list, client_prefix_list, job_list, address_list
import xlrd
from pyautogui import *
from time import sleep



class TransactionSegment(object):
    def __init__(self, packet):
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
        self.link = self.create_link()

    def find_location(self):
        for record in job_list:
            if self.str_format(record[0]) == self.job_in:
                return str(record[3])

    def find_program_date(self):
        for record in job_list:
            if self.str_format(record[0]) == self.job_in:
                return self.format_date(record[2])

    def create_link(self):
        directory = r'\\crgnas01\shared\CRG Internal Files\Financial Files\Support docs\YR2016\Mar'
        doc = '\%s' % (self.doc_name)
        link = directory + doc
        return link

    def str_format(self, input):
        string = str(input)
        if string.find('.') > 0:
            string = string.strip('.0')
        return string

    def format_date(self, date):
        if len(str(date)) == 0:
            return None
        elif len(str(date)) == 8:
            return date
        else:
            try:
                year, month, day, hour, minute, second = xlrd.xldate_as_tuple(date, 0)
                formatted_date = str(month) + '/' + str(day) + '/' + str(year)[2:]
                return formatted_date
            except:
                return date

    def bold(self, string):
        return '\033[1m' + string + '\033[0m'

    def clean_ref_num(self, raw_ref_num):
        str_ref_num = str(raw_ref_num)
        if str_ref_num[-2:] == '.0':
            formatted_ref_num = str_ref_num[:len(str_ref_num) - 2]
            return formatted_ref_num
        else:
            return str_ref_num



class LineItem(TransactionSegment):
    def __init__(self, packet):
        TransactionSegment.__init__(self,packet)
        self.vendor_obj = self.create_vendor_obj()
        self.full_job = self.find_full_job_name()
        self.short_job = self.find_short_job()
        self.job_out = self.create_job_out()
        self.client = self.find_client()
        self.gl_out = self.create_gl_out()
        self.program_date = self.find_program_date()
        self.memo = self.create_memo()
        self.tran_class = self.create_trans_class()
        self.tabulate_box = [self.gl_out,str(self.amount),self.memo,self.job_out]

    def find_full_job_name(self):
        for record in job_list:
            if self.str_format(record[0]) == self.job_in:
                return record[1]

    def find_short_job(self):
        for record in job_list:
            if self.str_format(record[0]) == self.job_in:
                return record[4]

    def find_client(self):
        if self.full_job:
            for record in client_prefix_list:
                if record[0] == self.full_job[:3]:
                    return record[1]

    def create_gl_out(self):
        '''
        This method returns a string that is the final input into QB
        EX: "Meeting Costs:Fees:Cynapsus"
        '''
        for record in gl_list:
            if record[0] == self.gl_in:
                if self.client:
                    return record[1] + self.client    #for COGS accounts
                else:
                    return record[1]      #for non-COGS accounts

    def create_memo(self):
        '''
        This method returns a string that becomes the memo line of a line item.
        '''
        if self.vendor_obj:
            loc = self.find_location()
            return self.vendor_obj.first_name + ' ' + self.vendor_obj.last_name + ' - ' + self.comments + ' - ' + str(loc) + ' - ' + str(self.program_date)
        else:
            return 'ERROR'


    def create_vendor_obj(self):
        '''
        This method takes the vendor_name and cross-references it against the vendor database. Outputs vendor obj
        '''
        for record in address_list:
            if record[0] == self.vendor_name:
                obj = Vendor(record)
                return obj

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

    def create_job_out(self):
        '''
        Outputs string for full job name.
        EX: "CYN6002 National Scientific Ad Board:Ground"
        '''
        if self.full_job:
            return str(self.full_job) + str(':') + self.gl_in
        else: return None

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
        quick = locateOnScreen('img/quick.png')
        quick2 = locateOnScreen('img/quick2.png')
        if quick or quick2:
            press('q')
        sleep(1.5)
        press('tab')
        if self.tran_class:
            typewrite(self.tran_class)
        sleep(1)
        press('tab')
        sleep(1.5)
        quick = locateOnScreen('img/quick.png')
        quick2 = locateOnScreen('img/quick2.png')
        if quick or quick2:
            press('q')
        sleep(1.5)

    def __str__(self):
        print '______________line item______________'
        print self.bold('General ledger: ') + ' %s' % (self.gl_out)
        print self.bold('Amount        : ') + ' %s' % (self.amount)
        print self.bold('Memo          : ') + ' %s' % (self.memo)
        print self.bold('Job           : ') + ' %s' % (self.job_out)
        print self.bold('Class         : ') + ' %s' % (self.tran_class)

class LineItemInv(LineItem):
    def __init__(self, packet):
        LineItem.__init__(self, packet)

    def create_memo(self):
        if not self.vendor_obj:
            pass
        elif self.comments:
            return self.vendor_obj.payable + ' - ' + self.gl_in + ' - ' + self.comments
        else:
            return self.vendor_obj.payable + ' - ' + self.gl_in


class ContainerTransaction(TransactionSegment):
    def __init__(self, packet):
        TransactionSegment.__init__(self, packet)
        self.line_items = []
        self.i = 0

    def next(self):
        if self.i == len(self.line_items):
            raise StopIteration
        current_item = self.line_items[self.i]
        self.i += 1
        return current_item

    def __iter__(self):
        return self

    def len(self):
        return len(self.line_items)

class Invoice(ContainerTransaction):
    def __init__(self, packet):
        ContainerTransaction.__init__(self,packet)
        self.tabulate_box = [self.vendor_name, self.date, self.ref_num,str(self.amount)]

    def append(self, line_item):
        self.line_items.append(line_item)

    def QB(self):
        typewrite(self.vendor_name)
        press('tab')
        sleep(1)
        typewrite(self.date)
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
        print self.bold('Vender        : ') + ' %s' % (self.vendor_name)
        print self.bold('Date          : ') + ' %s' % (self.date)
        print self.bold('Invoice Num   : ') + ' %s' % (self.ref_num)
        print self.bold('Amount        : ') + ' %s' % (self.amount)
        print self.bold('Support doc   : ') + ' %s' % (self.link)


class CheckRequest(ContainerTransaction):
    def __init__(self, packet):
        ContainerTransaction.__init__(self,packet)
        self.vendor_obj = self.create_vendor_obj()
        self.memo = self.create_memo()
        self.tabulate_box = [self.vendor_name, self.date, self.memo,str(self.amount)]

    def append(self, line_item):
        self.line_items.append(line_item)

    def create_vendor_obj(self):
        for record in address_list:
            if record[0] == self.vendor_name:
                obj = Vendor(record)
                return obj

    def create_memo(self):
        if self.comments == 'Fee For Service':
            checktype = 'Fee For Service'
        else:
            checktype = 'Expenses'
        name = self.vendor_obj.last_name
        location = self.find_location()
        program_date = self.find_program_date()
        memo = name + ' - ' + checktype + ' - ' + str(location) + ' - ' + str(program_date)
        return memo

    def QB(self):
        typewrite(self.vendor_name)
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
        typewrite(self.link)
        sleep(3)
        hotkey('alt','s')

    def __str__(self):
        print '______________TRANSACTION TOP______________'
        print self.bold('Vender        : ') + ' %s' % (self.vendor_name)
        print self.bold('Amount        : ') + ' %s' % (self.amount)
        print self.bold('Memo          : ') + ' %s' % (self.memo)
        print self.bold('Support doc   : ') + ' %s' % (self.link)


class Vendor(object):
    def __init__(self, packet):
        self.vendor = packet[0]
        self.tax_ID = packet[2]
        self.address1 = packet[2]
        self.address2 = packet[6]
        self.city = packet[8]
        self.state = packet[10]
        self.zip = packet[12]
        self.last_name = packet[14]
        self.first_name = packet[16]
        self.salutation = packet[18]
        self.payable = packet[28]
