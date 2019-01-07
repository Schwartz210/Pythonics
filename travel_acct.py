'''
The purpose of this scrip is to pull the data from the travel account and output in a special format for review
'''
from spreadsheet import format_one_job_number, make_spreadsheet, PowerRead
from folder_contents import amex_filepath
from super_transaction import SuperTransaction
from logger import logger

class Trans(SuperTransaction):
    '''
    Trans class instances contain basic info about a transaction
    '''
    def __init__(self, record):
        self.record= record
        self.vendor = record[0]
        self.date = record[1]
        self.amount = record[2]
        self.num = record[3]
        self.travel_date = record[4]
        self.comment = record[5]
        self.code = record[6]
        self.job_in = record[7]
        self.person = record[8]
        self.split = record[9]
        self.booking = self.set_booking(record)

    def set_booking(self, record):
        '''
        Sets self.booking either to 1) basic job ('ABB-N7272'), 2) Non-COGS acct ('NBD Mtg.'), or 3) multiple jobs.
        '''
        if self.is_standard_single_job():
            return self.find_short_job(format_one_job_number(self.job_in))
        elif self.is_non_COGS():
            return self.code
        elif self.is_split():
            return self.split_handler(record)
        else:
            return 'Needs work'

    def is_standard_single_job(self):
        '''
        Evaluates if transaction has just one job. Returns boolean
        '''
        if self.split == '' and self.job_in != '-':
            return True
        else:
            return False

    def is_non_COGS(self):
        '''
        Evaluates if transaction has no job but rather is non_COGS. Returns boolean
        '''
        if self.job_in == '-' and self.split == '':
            return True
        else:
            return False

    def is_split(self):
        '''
        Evaluates if transaction is split. Returns boolean
        '''
        if self.split != '':
            if self.split[0:4] == 'Jobs':
                return True
            elif self.split[0:2] == 'GL':
                return True
            else:
                raise Exception('Cannot categorize:', self.record)
        else:
            return False

    def split_handler(self, record):
        '''
        Returns multiple short jobs as String
        '''
        out = ''
        if self.split[0:4] == 'Jobs':
            jobs = self.split[5:].split(',')
            for job in jobs:
                out += str(self.find_short_job(job)) + ','
        elif self.split[0:2] == 'GL':
            ledgers = self.split[3:].split(',')
            for ledger in ledgers:
                out += ledger + ','
        else:
            raise Exception('Some other typer of split I am unprepared to handle.')

        try:
            out = out[:-1]
        except:
            raise Exception('No splits found in this record:', record)
        return out

    def to_array(self):
        return [self.vendor, self.date, self.amount, self.num, self.travel_date, self.comment, self.booking, self.person]


@logger
def execute_make_travel_report(amex):
    '''
    Executes entire process
    '''
    file = amex_filepath(amex) + '.xlsm'
    power = PowerRead(file,
                      'Transactions',
                      columns=['B','C','E','F','I','J','K','L','M','O'],
                      factory=Trans,
                      only_include={'A':'Geoge Smith'})
    rows = power.inventory_to_array()
    make_spreadsheet(rows,
                     title=r'H:\Monthly travel report.xls',
                     headers=['Vendor','Date','Amount','#','Travel Date','Comment','Booking','Person'])

