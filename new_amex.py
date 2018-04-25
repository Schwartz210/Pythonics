import csv
from spreadsheet import open_file
from publish import publish_to_template
from folder_contents import open_file2
from logger import logger

class Transaction(object):
    def __init__(self, packet, trans_num_dict, vendor_dict, airport_dict):
        self.packet = packet
        self.trans_num_dict = trans_num_dict
        self.vendor_dict = vendor_dict
        self.airport_dict = airport_dict
        self.account = packet[3]
        self.vendor_in = packet[2] + packet[11]
        self.vendor = None
        self.amex_date = packet[0][:10]
        self.amount = float(packet[7])
        self.travel_date = None
        self.airport0 = None
        self.airport1 = None
        self.airport2 = None
        self.airport3 = None
        self.airport4 = None
        self.airports = []
        self.airfare_name = None
        self.comments = ''
        self.trans_number = self.packet[14].strip("'")
        self.set_travel_info()
        self.set_comments()
        self.set_vendor()
        self.template_col_dict = {'A' : self.account,
                                  'B' : self.vendor,
                                  'C' : self.amex_date,
                                  'D' : '',     #INTENTIONALLY LEFT BLANK
                                  'E' : self.amount,
                                  'F' : '',     #INTENTIONALLY LEFT BLANK
                                  'G' : '',     #INTENTIONALLY LEFT BLANK
                                  'H' : '',     #INTENTIONALLY LEFT BLANK
                                  'I' : self.travel_date,
                                  'J' : self.comments,
                                  'K' : '',     #INTENTIONALLY LEFT BLANK
                                  'L' : '',     #INTENTIONALLY LEFT BLANK
                                  'M' : self.airfare_name,
                                  'N' : '',     #INTENTIONALLY LEFT BLANK
                                  'O' : ''}     #INTENTIONALLY LEFT BLANK

    def set_travel_info(self):
        if self.trans_number in self.trans_num_dict:
            self.airport0 = self.trans_num_dict[self.trans_number][0]
            self.airport1 = self.trans_num_dict[self.trans_number][1]
            self.airport2 = self.trans_num_dict[self.trans_number][2]
            self.airport3 = self.trans_num_dict[self.trans_number][3]
            self.airport4 = self.trans_num_dict[self.trans_number][4]
            self.airfare_name = self.trans_num_dict[self.trans_number][5]
            self.travel_date = self.fix_date(self.trans_num_dict[self.trans_number][6])
            self.airports = [self.airport0, self.airport1, self.airport2, self.airport3, self.airport4]

    def set_comments(self):
        if self.airport0 and not self.airport0 == 'N/A':
            self.comments = '#'
            for airport in self.airports:
                if airport and not airport == 'N/A' and airport in self.airport_dict:
                    self.comments += self.airport_dict[airport] + '->'
            self.comments = self.comments[0:len(self.comments)-2]

    def __str__(self):
        return self.comments

    def fix_date(self, date_in):
        formatted_date = date_in[:2] + '/' + date_in[2:] + '/2018'
        if formatted_date.find('//') > -1:
            return None
        return formatted_date

    def set_vendor(self):
        if self.vendor_in in self.vendor_dict:
            self.vendor = self.vendor_dict[self.vendor_in]
        else:
            pass


def make_trans_num_dictionary(ofx_data):
    ofx_dict = {}
    for record in ofx_data:
        if record[8] != '':
            ofx_dict[record[1]] = [record[8], record[11], record[14], record[17], record[20], record[24], record[25]]
    return ofx_dict

def get_data(file):
    data = []
    with open(file,'r') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        for record in reader:
            if len(record) > 0:
                data.append(record)
    return data

def clean_ofx(data):
    for record in data:
        record[1] = record[1].strip('Reference: ')
        if len(record) < 25:
            for _ in range(25-len(record)):
                record.append('')
    return data

def get_vendor_dictionary():
    vendor_data = open_file(r'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\Arrays\Vendor name array.xlsm',0,10000)
    vendor_dict = {}
    for record in vendor_data:
        vendor_dict[record[0]] = record[1]
    return vendor_dict

def get_airport_dictionary():
    airport_data = open_file(r'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\Arrays\Location array.xlsx',0,300)
    airport_dict = {}
    for record in airport_data:
        airport_dict[record[0]] = record[1]
    return airport_dict

def print_new_vendors(transactions, vendor_dict):
    new_vendors = set([])
    for transaction in transactions:
        if transaction.vendor_in not in vendor_dict:
            new_vendors.add(transaction.vendor_in)
    print('New Vendors:')
    for vendor in new_vendors:
        print(vendor)

def print_new_airports(transactions, airport_dict):
    new_airports = set([])
    for transaction in transactions:
        for airport in transaction.airports:
            if airport not in airport_dict:
                new_airports.add(airport)
    print('\n\n' + 'New Airports:')
    for airport in new_airports:
        print(airport)

@logger
def new_amex_execute(publish=False):
    ofx_data = get_data('H:\ofx.csv')
    ofx_data_clean = clean_ofx(ofx_data)
    trans_num_dict = make_trans_num_dictionary(ofx_data_clean)
    transaction_data = get_data('H:\Transactions.csv')
    transactions = []
    vendor_dict = get_vendor_dictionary()
    airport_dict = get_airport_dictionary()
    for record in transaction_data:
        transactions.append(Transaction(record, trans_num_dict, vendor_dict, airport_dict))
    if publish:
        template = r'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\Amex Template.xlsm'
        file_destination = r'H:\New Amex.xlsm'
        publish_to_template(transactions,file_destination,template,'Transactions')
    else:
        print_new_vendors(transactions, vendor_dict)
        print_new_airports(transactions, airport_dict)
        open_file2('vendor')
        open_file2('airport')




