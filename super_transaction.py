__author__ = 'aschwartz - Schwartz210@gmail.com'
from xlrd import xldate_as_tuple
from datetime import datetime
from spreadsheet import open_file
from folder_contents import financial
from mydatabases import names, jobs, clients, general_ledger

QUICKBOOKS_CLOSING_DATE = datetime.strptime('06/01/2016', '%m/%d/%Y')


class SuperTransaction(object):
    '''
    This class is not an actual transaction. It has methods associated with various types of transactions.
    This class is designed purely to pass inheritance to child classes, which actually represent transactions.
    The methods here are designed for the broadest possilbe useages.
    '''
    def __init__(self):
        pass

    def str_format(self, input):
        string = str(input)
        if string.find('.') > 0:
            string = string.strip('.0')
        return string

    def format_date(self, date):
        if len(str(date)) == 0:
            return None
        else:
            try:
                year, month, day, hour, minute, second = xldate_as_tuple(date, 0)
                formatted_date = str(month) + '/' + str(day) + '/' + str(year)
                return formatted_date
            except:
                return date

    def find_person(self, person):
        '''
        Returns a person's full name, and QB class based on Amex name.
        '''
        for record in names:
            if record[2] == person:
                return record[0], record[1]
        return 'ERROR', 'ERROR'

    def find_person_class(self, person):
        '''
        Returns a person's QB class based on full name
        '''
        if person == None: return None
        for record in names:
            if record[0] == person:
                return record[1]
        return 'ERROR'

    def find_full_job_name(self, job_in):
        '''
        Converts short job name --> long job name
        EXAMPLE: "P6189" --> "KEM-P6189 Job name"
        '''
        if job_in == '-':
            return None
        for record in jobs:
            if self.str_format(record[0]) == self.str_format(job_in):
                return record[1]
        return None

    def find_short_job(self, job_in):
        for record in jobs:
            if self.str_format(record[0]) == job_in:
                return record[4]

    def find_client(self, full_job):
        '''
        Determines which client a job is associated with based on prefix.
        '''
        if full_job:
            for record in clients:
                if record[0] == full_job[:3]:
                    return record[1]

    def create_gl_out(self, gl_in, client):
        '''
        This method returns a string that is the final input into QB
        EXAMPLE: 'Fees' and 'XYZ Pharma' --> 'Meeting Costs:Fees:XYZ Pharma'
        '''
        for record in general_ledger:
            if record[0] == gl_in:
                if client:
                    return record[1] + client    #for COGS accounts
                else:
                    return record[1]      #for non-COGS accounts

    def create_job_out(self, full_job, gl_in):
        '''
        Outputs string for full job name.
        EX: "CYN6002 National Scientific Ad Board:Ground"
        '''
        if full_job:
            try: return str(full_job) + str(':') + gl_in
            except: print "There is an issue with this job: ", full_job
        else:
            return None

    def create_link(self, directory, doc_name):
        '''
        Outsputs a link to a file with a filepath.
        '''
        if doc_name:
            doc = '\%s' % (doc_name)
            link = directory + doc
            return link
        else:
            return None

    def find_quickbooks_date(self, date):
        '''
        Any dates prior to QUICKBOOKS_CLOSING_DATE, become QUICKBOOKS_CLOSING_DATE.
        Any dates on or after QUICKBOOKS_CLOSING_DATE are passed through.
        '''
        d2 = datetime.strptime(date,'%m/%d/%Y')
        difference = (d2 - QUICKBOOKS_CLOSING_DATE).days
        if difference < 0:
            return QUICKBOOKS_CLOSING_DATE.strftime('%m/%d/%Y')
        else:
            return date

    def null_clean(self, input):
        '''
        Converts null values to boolean None.
        '''
        if input == '':
            return None
        else:
            return input

    def create_vendor_obj(self):
        '''
        Factory for instances of 'Vendor' class
        '''
        filepath = financial + r'\Avi\Address.xlsx'
        data = open_file(filepath,1,15000)
        all_names = []
        for record in data:
            all_names.append(record[0])
        if self.vendor_name in all_names:
            idx = all_names.index(self.vendor_name)
            obj = Vendor(data[idx])
            return obj
        else:
            return 'ERROR: Record not found'

class Vendor(object):
    '''
    This class is designed to receive a vendor record from a QuickBooks export file, and map it to fields.
    '''
    def __init__(self, packet):
        self.vendor = packet[0]
        self.tax_ID = packet[2]
        self.address1 = packet[4]
        self.address2 = packet[6]
        self.city = packet[8]
        self.state = packet[10]
        self.zip = packet[12]
        self.last_name = packet[14]
        self.first_name = packet[16]
        self.salutation = packet[18]
        self.payable = self.fix_payable(packet[28])

    def fix_payable(self, payable_name):
        if payable_name == '':
            return self.vendor
        return payable_name


