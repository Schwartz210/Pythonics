__author__ = 'aschwartz - Schwartz210@gmail.com'
from xlrd import xldate_as_tuple
from datetime import datetime
from spreadsheet import open_file
from folder_contents import financial
from sql_time import pull_data, exists

QUICKBOOKS_CLOSING_DATE = datetime.strptime('03/01/2018', '%m/%d/%Y')


class SuperTransaction(object):
    '''
    This class is not an actual transaction. It has methods associated with various types of transactions.
    This class is designed purely to pass inheritance to child classes, which actually represent transactions.
    The methods here are designed for the broadest possible useages.
    '''
    def __init__(self):
        self.missing_job = False

    def str_format(self, input):
        string = str(input)
        if string.find('.') > 0:
            string = string.strip('.0')
        more_zeros = 4 - len(string)
        for _ in range(more_zeros):
            string += '0'
        if string == '0000':
            return None
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
        sql_request = 'SELECT Full, Quickbooks FROM names WHERE Airline="%s"' % (person)
        if exists(sql_request):
            data = pull_data(sql_request)
            return data[0]
        raise Exception('Add %s to the Name-Class Array' % (person))

    def find_person_class(self, person):
        '''
        Returns a person's QB class based on full name
        '''
        sql_request = 'SELECT quickbooks FROM names WHERE full="%s"' % (person)
        if exists(sql_request):
            data = pull_data(sql_request)
            return data[0][0]
        return 'ERROR'

    def find_full_job_name(self, job_in):
        '''
        Converts short job name --> long job name
        EXAMPLE: "P6189" --> "KEM-P6189 Norco bioequivalence"
        '''
        if job_in == '-': return None
        sql_request = 'SELECT full FROM jobs WHERE number="%s"' % (job_in)
        if exists(sql_request):
            data = pull_data(sql_request)
            return data[0][0]
        else:
            return None

    def find_full_job_name_from_short(self, short_job):
        '''
        Converts short job name --> long job name
        EXAMPLE: "KEM-P6189" --> "KEM-P6189 Norco bioequivalence"
        '''
        if short_job == '-': return None
        sql_request = 'SELECT full FROM jobs WHERE Short="%s"' % (short_job)
        if exists(sql_request):
            data = pull_data(sql_request)
            return data[0][0]
        else:
            return None

    def find_short_job(self, job_in):
        sql_request = 'SELECT short FROM jobs WHERE number="%s"' % (job_in)
        if exists(sql_request):
            data = pull_data(sql_request)
            return data[0][0]
        else:
            self.missing_job = job_in
            return None

    def find_client(self, full_job):
        '''
        Determines which client a job is associated with based on prefix.
        '''
        if full_job:
            sql_request = 'SELECT full FROM clients WHERE short="%s"' % (full_job[:3])
            if exists(sql_request):
                data = pull_data(sql_request)
                return data[0][0]

    def create_gl_out(self, gl_in, client):
        '''
        This method returns a string that is the final input into QB
        EXAMPLE: 'Fees' and 'XYZ Pharma' --> 'Meeting Costs:Fees:XYZ Pharma'
        '''
        if gl_in == '':
            return None
        sql_request = 'SELECT full FROM general_ledger WHERE short="%s"' % (gl_in)
        data = pull_data(sql_request)
        if not data:
            raise Exception('GL issue with ', gl_in)
        gl_full = data[0][0]
        if client:
            return gl_full + client
        else:
            return gl_full

    def create_job_out(self, full_job, gl_in):
        '''
        Outputs string for full job name.
        EX: "CYN6002 National Scientific Ad Board:Ground"
        '''
        if full_job:
            try:
                return str(full_job) + str(':') + str(gl_in)
            except:
                print "There may be an issue with this job: ", full_job
                print "Or there may be issue with this gl: ", gl_in
                print 'Or it may be the fault of a combo: ', str(full_job) + str(':') + str(gl_in)
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
        data = open_file(filepath,1,16000)
        all_names = []
        for record in data:
            all_names.append(record[0])
        if self.vendor_name in all_names:
            idx = all_names.index(self.vendor_name)
            obj = Vendor(data[idx])
            return obj
        else:
            raise Exception('ERROR: Record not found: ', self.vendor_name)


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


