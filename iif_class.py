from spreadsheet import open_file, screen_out_col, screen_out_empty_rows, format_one_job_number, format_date
from super_transaction import SuperTransaction, QUICKBOOKS_CLOSING_DATE
from csv import writer, QUOTE_MINIMAL
from window import SingleSelectionWindow, MultiSelectionWindow
from folder_contents import amex_support, get_folder_contents, clean, amex_filepath
from QBnames import main_inputter_subroutine
from startup import Startup
from time import sleep
from pyautogui import typewrite, press, hotkey
from transplitr import transplitr, split_based_on_jobs, split_based_on_GLs
from sql_time import pull_data
from attendee_report import get_line_items_from_attendee_reports
from logger import logger
__author__ = 'aschwartz - Schwartz210@gmail.com'

class IIFTransaction(SuperTransaction):
    def __init__(self):
        pass

    def create_TRNSTYPE(self):
        '''
        This is a field in the Intuit Interchange File template.
        Returns 'CREDIT CARD' for positive amounts, and 'CCARD REFUND' for negative.
        '''
        if self.amount >= 0:
            return 'CREDIT CARD'
        else:
            return 'CCARD REFUND'

    def cc_acct(self):
        sql_request = 'SELECT full FROM cc_accts WHERE short="%s"' % (self.account)
        data = pull_data(sql_request)
        try:
            return data[0][0]
        except:
            raise Exception('CC Acct error: ', sql_request)

    def create_memo(self):
        '''
        String formatting for memo field.
        '''
        memo = self.vendor + ' - ' + self.gl_in
        optional_fields = [self.person_full_name, self.comments, self.CDate]
        for field in optional_fields:
            if field:
                memo += ' - ' + field
        return memo

    def create_iff(self):
        '''
        This method converts the basic transaction information into a list of IIF formatted rows.
        '''
        output_list = []
        TRNS = ['TRNS','',self.TRNSTYPE,self.qb_date,self.ACCNT, self.vendor,self.amount * -1, self.ref_no, '',self.link]
        output_list.append(TRNS)
        if self.split:
            lines = self.split_handler()
            for SPL in lines:
                output_list.append(SPL)
        else:
            SPL = ['SPL','',self.TRNSTYPE,self.qb_date,self.gl_out,self.job_out, self.amount,self.ref_no,self.person_class,self.memo]
            output_list.append(SPL)
        ENDTRNS = ['ENDTRNS']
        output_list.append(ENDTRNS)
        return output_list

    def create_directory(self):
        amex = self.ref_no[:7]
        directory = amex_support(amex)
        return directory

    def get_line_items(self):
        '''
        Returns line_items based on split_type. This may involve pulling data from separate spreadsheets, tabs on the-
        same spreadsheet, job or GL cues listed in the split field (Column O).
        '''
        if self.CDate:
            date = self.CDate
        else:
            date = self.amex_date
        if self.split.find('XLSX:') > -1:
            return get_line_items_from_attendee_reports(self.create_link(self.directory, self.split[5:]))
        elif self.split.find('XLS:') > -1:
            return self.split_pull(self.create_link(self.directory, self.split[4:]))
        elif self.split.find('Tab:') > -1:
            return transplitr(self.amex, self.split[4:], self.amount, self.vendor)
        elif self.split.find('Jobs:') > -1:
            return split_based_on_jobs(self.vendor, self.gl_in, self.split, self.person_full_name, self.amount, date, self.comments)
        elif self.split.find('GLS:') > -1:
            return split_based_on_GLs(self.vendor, self.split, self.person_full_name, self.amount, date, self.comments)
        else:
            raise Exception('No split_type recognized: ', self.split)


    def split_handler(self):
        '''
        This method is triggered during the self.create_iff() process for those items with splits. Splits are-
        transactions with more than one line items. Split data can be stored on other spreadsheets.
        This methods pulls are converts this data into IIF format.
        '''
        line_items = self.get_line_items()
        output_list = []
        self.split_gl_out = set()
        self.split_job_out = set()
        for line_item in line_items:
            SPL = ['SPL','',self.TRNSTYPE,self.qb_date,line_item[0],line_item[3], line_item[1],self.ref_no,line_item[4],line_item[2]]
            output_list.append(SPL)
            self.split_gl_out.add(line_item[0])
            self.split_job_out.add(line_item[3])
        return output_list

    def split_pull(self, database):
        '''
        This function pulls line items from spreadsheets that contain transaction line items.
        '''
        tab = 0
        start_column = 0
        end_column = start_column + 5
        data = open_file(database,tab,200)
        data = screen_out_col(data, start_column, end_column)
        output_data = screen_out_empty_rows(data)
        del output_data[0]
        for row in output_data:
            if row[4] == 0:
                row[4] = ''
        return output_data

class Factory(object):
    '''
    This is the parent class for the two factory classes 'PayrollFactory' and 'AmexFactory'
    '''
    def __init__(self):
        self.output_filename = ''

    def text_to_columns(self):
        '''
        When the file is first produced all contents are in column A, separated by semicolons.This method opens up the-
        new .iff and delimits the text into columns.
        This method directly takes over keyboard using pyautogui package.
        '''
        sleep(5)
        Startup.excel()                   #opens excel
        sleep(3)
        hotkey('ctrl','o')                #'Open' file dialogue
        sleep(1)
        press('right')
        for _ in range(3):
            press('down')
            sleep(.25)
        press('esc')
        for _ in range(5):
            press('down')
        sleep(1)
        press('enter')
        sleep(1)
        typewrite('H:\Amex.iif')   #types file name
        press('enter')
        sleep(1)
        press('d')                        #selects 'Delimited' option
        sleep(.5)
        press('n')                        #selects 'Next' option
        sleep(.5)
        press('m')                        #selects 'Semicolon' option
        sleep(.5)
        press('f')                        #selects 'Finish' option
        sleep(.5)
        hotkey('ctrl','s')                #Saves file
        sleep(.5)
        press('y')                        #selects 'Yes' option
        hotkey('alt','f')                 #selects 'Office Button' in Excel
        press('x')                        #exits Excel
        press('n')                        #selects 'No' option
        sleep(1)


class PayrollLine(IIFTransaction):
    '''
    This class represent individual lines items in the payroll allocation report. It serves to produce a single row-
    on the payroll journal (Intuit Interchange File) for each line item.
    '''
    def __init__(self, packet):
        self.job_in = packet[0]
        self.client_ver = packet[1]
        self.amount = packet[1]
        self.gl_in = 'Labor'
        self.full_job = self.find_full_job_name(self.job_in)
        self.job_out = self.create_job_out(self.full_job, self.gl_in)
        self.client = self.find_client(self.full_job)
        self.gl_out = self.create_gl_out(self.gl_in, self.client)

    def iif(self):
        '''
        This method produces the row for the .iff to be produced.
        '''
        row = ['SPL',None,'GENERAL JOURNAL',self.date,self.gl_out,None,self.amount,self.DOCNUM, self.memo, self.job_out]
        return row

    def __str__(self):
        return ['Line Item: ', self.gl_out,self.amount,self.job_out]


class PayrollFactory(Factory):
    '''
    This class functions to produce an Intuit Interchange File. It produces many 'PayrollLine' instances and extracts-
    their outputs, which get assembled onto a report.
    '''
    def __init__(self):
        self.data = self.payroll_pull()
        self.memo = 'Job Labor - December 2016'
        self.DOCNUM = 'PR 12/16'
        self.date = '12/31/2016'
        self.gl_out = 'Payroll Allocation of Job Labor'
        self.output_filename = r"H:\Amex.iif"
        self.line_items = self.factory_method()
        self.modify_line_items()
        self.amount = self.get_total()
        self.iff_info = self.create_iif_rows()
        self.__str__()
        self.publish()
        #self.text_to_columns()

    def payroll_pull(self):
        '''
        Pulls data from payroll file. Returns list.
        '''
        filepath = 'S:\CRG Internal Files\Financial Files\Payroll allocation\Source files\DecemberLabor.xlsx'
        data = open_file(filepath, 0, 100)
        del data[:10]
        processed_data = []
        for row in data:
            if row[0] != '':
                new_row = [format_one_job_number(row[0]),row[8]]
                processed_data.append(new_row)
        return processed_data

    def factory_method(self):
        '''
        This method produces instances of 'PayrollLine' class.
        '''
        output_list = []
        for row in self.data:
            line = PayrollLine(row)
            output_list.append(line)
        return output_list

    def modify_line_items(self):
        '''
        This method shares self.attributes with all instances of 'PayrollLine' class.
        '''
        for line_item in self.line_items:
            line_item.memo = self.memo
            line_item.DOCNUM = self.DOCNUM
            line_item.date = self.date

    def get_total(self):
        '''
        This method sums up the amounts of all instances of 'PayrollLine' class to add to the report.
        '''
        total = 0.0
        for line_item in self.line_items:
            total += line_item.amount
        return round(total,2) * -1

    def create_iif_rows(self):
        '''
        This method gets all the iif. rows and appends them to a list.
        '''
        output_list = []
        output_list.append(['!TRNS','TRNSID','TRNSTYPE','DATE','ACCNT','CLASS','AMOUNT','DOCNUM','MEMO','NAME'])
        output_list.append(['!SPL','SPLID','TRNSTYPE','DATE','ACCNT','CLASS','AMOUNT','DOCNUM','MEMO','NAME'])
        output_list.append(['!ENDTRNS'])
        output_list.append(['TRNS',None,'GENERAL JOURNAL',self.date,self.gl_out,None,self.amount,self.DOCNUM,self.memo,None])
        for line_item in self.line_items:
            iif_row = line_item.iif()
            output_list.append(iif_row)
        output_list.append(['ENDTRNS'])
        return output_list

    def publish(self):
        '''
        Creates Intuit Interchange File(.iif)
        '''
        with open(self.output_filename, 'w') as csvfile:
            filewriter = writer(csvfile, delimiter=';', quotechar='|', quoting=QUOTE_MINIMAL, lineterminator="\n")
            for row in self.iff_info:
                try:
                    filewriter.writerow(row)
                except:
                    print row

    def __str__(self):
        for row in self.iff_info:
            print row


class AmexTransaction(IIFTransaction):
    def __init__(self, packet, amex):
        self.account = packet[0]
        self.vendor = packet[1]
        self.amex_date = self.format_date(packet[2])
        self.ref_no = packet[3]
        self.amount = packet[4]
        self.num = packet[5]
        self.receipt = packet[6]
        self.CDate = format_date(packet[8])
        self.comments = packet[9]
        self.gl_in = packet[10]
        self.job_in = format_one_job_number(packet[11])
        self.name_in = packet[12]
        self.amex = amex
        self.split = self.null_clean(packet[14])
        self.qb_date = self.find_quickbooks_date(self.amex_date)
        self.TRNSTYPE = self.create_TRNSTYPE()
        self.ACCNT = self.cc_acct()
        if self.name_in:
            self.person_full_name, self.person_class = self.find_person(self.name_in)
        else:
            self.person_full_name = None
            self.person_class = None
        self.memo = self.create_memo()
        self.full_job = self.find_full_job_name(self.job_in)
        self.job_out = self.create_job_out(self.full_job, self.gl_in)
        self.client = self.find_client(self.full_job)
        self.gl_out = self.create_gl_out(self.gl_in, self.client)
        self.directory = self.create_directory()
        self.link = self.create_link(self.directory, self.receipt)


class AmexFactory(Factory):
    def __init__(self):
        self.amex = self.select_amex()
        self.data = self.amex_pull(self.amex)
        self.all_accts = self.find_accts(self.data)
        self.accts = self.pick_accts()
        self.transactions = self.create_transactions()
        self.output_filename = r"H:\Amex.iif"
        self.i = 0
        self.iff_info = []
        self.create_iif_headers()
        self.convert_trans_to_iif()
        self.publish()
        self.__str__()
        #self.text_to_columns()
        self.qb_inputter()
        #open_file2('amex_prog')

    def get_options(self):
        '''
        Retrieves, cleans, removes duplicates, sorts, platinum and centurion folders' files.
        Output is used to populat options memu for AmexFactory
        '''
        platinum_files = get_folder_contents('plat')
        centurion_files = get_folder_contents('cent')
        all_files = clean(platinum_files + centurion_files)
        all_files2 = set([])
        for file in all_files:
            file_without_suffix = file[:7]
            all_files2.add(file_without_suffix)
        all_files3 = list(all_files2)
        all_files3.sort(reverse=True)
        return all_files3[:15]

    def select_amex(self):
        '''
        Opens window where user selects which Amex spreadsheet with provide source data.
        '''
        options = self.get_options()
        window = SingleSelectionWindow(options)
        selected_option = window.selected_options[0]
        selected_option = amex_filepath(selected_option)
        selected_option = str(selected_option) + '.xlsm'
        return selected_option

    def amex_pull(self, src):
        amex_data = open_file(src,0,1000)
        del amex_data[0]
        return amex_data

    def find_accts(self, data):
        result = set()
        for row in data:
            if row[0] != 'Main':
                result.add(row[0])
        return result

    def pick_accts(self):
        '''
        Opens up GUI window and user selects accounts
        '''
        window = MultiSelectionWindow(self.all_accts)
        return window.selected_options

    def create_transactions(self):
        '''
        AmexTransaction object factory method. Produces lots of these buggers
        '''
        output_list = []
        for row in self.data:
            if row[0] in self.accts:
                transaction = AmexTransaction(row, self.amex)
                output_list.append(transaction)
        return output_list

    def next(self):
        if self.i == len(self.transactions):
            raise StopIteration
        current_item = self.transactions[self.i]
        self.i += 1
        return current_item

    def __iter__(self):
        return self

    def create_iif_headers(self):
        self.iff_info.append(['!TRNS','TRNSID','TRNSTYPE','DATE','ACCNT','NAME','AMOUNT','DOCNUM','CLASS','MEMO'])
        self.iff_info.append(['!SPL','SPLID','TRNSTYPE','DATE','ACCNT','NAME','AMOUNT','DOCNUM','CLASS','MEMO'])
        self.iff_info.append(['!ENDTRNS'])

    def convert_trans_to_iif(self):
        for transaction in self.transactions:
            iif_data = transaction.create_iff()
            for row in iif_data:
                self.iff_info.append(row)

    def publish(self):
        '''
        Creates Intuit Interchange File(.iif)
        '''
        with open(self.output_filename, 'w') as csvfile:
            filewriter = writer(csvfile, delimiter=';', quotechar='|', quoting=QUOTE_MINIMAL, lineterminator="\n")
            for row in self.iff_info:
                try:
                    filewriter.writerow(row)
                except:
                    print 'This row has an unprintable character: ', row

    def get_qb_name_info(self):
        '''
        Collects vendors, accounts, and jobs in order to pass it along to 'QBnames.main_inputter_subroutine()'
        '''
        vendors = set()
        gl = set()
        jobs = set()
        for transaction in self.transactions:
            vendors.add(transaction.vendor)
            if transaction.split:
                gl = gl.union(transaction.split_gl_out)
                jobs = jobs.union(transaction.split_job_out)
            else:
                gl.add(transaction.ACCNT)
                gl.add(transaction.gl_out)
                jobs.add(transaction.job_out)
        if None in gl: gl.remove(None)
        if None in jobs: jobs.remove(None)
        data = (vendors, gl, jobs)
        return data

    def qb_inputter(self):
        '''
        Opens up window associated with the feature which enters jobs and gl's into QuickBooks.
        '''
        options = ['Yes','No']
        window = SingleSelectionWindow(options)
        qb_name_info = self.get_qb_name_info()
        if window.selected_options[0] == 'Yes':
            main_inputter_subroutine(data=qb_name_info)

    def __str__(self):
        print '================================================================================'
        print 'Intuit Interchange File (.iif) has been created in the H Drive','\n'
        print 'Source file: ', self.amex
        print 'The following account(s) are in this file: ',
        for acct in self.accts:
            print '    ', acct
        print 'The current QuickBooks closing date is: ', QUICKBOOKS_CLOSING_DATE.strftime('%m/%d/%Y')
        print '================================================================================'

@logger
def amex_factory():
    AmexFactory()