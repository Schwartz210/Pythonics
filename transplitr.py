__author__ = 'aschwartz - Schwartz210@gmail.com'
from spreadsheet import open_file, format_date
from iif_class import IIFTransaction
from folder_contents import amex_filepath
import xlwt

class SPLT(IIFTransaction):
    def __init__(self, packet):
        self.gl_in = packet[0]
        self.job_in = packet[1]
        self.person = packet[2]
        self.amount = packet[3]
        self.date = format_date(packet[4])
        self.comments = packet[5]
        self.full_job = self.find_full_job_name(self.job_in)
        self.short_job = self.find_short_job(self.job_in)
        self.job_out = self.create_job_out(self.full_job, self.gl_in)
        self.client = self.find_client(self.full_job)
        self.gl_out = self.create_gl_out(self.gl_in, self.client)
        self.person_class = self.find_person_class(self.person)
        self.memo = self.create_memo()

    def create_memo(self):
        memo = vendor + ' - ' + self.gl_in
        optional_fields = [self.person, self.comments, self.date]
        for field in optional_fields:
            if field != '':
                memo += ' - ' + field

        return memo

def get_data(amex_doc, tab):
    file = amex_filepath(amex_doc) + '.xlsm'
    raw_data = open_file(file, tab, 100)
    del raw_data[0]
    data = []
    for row in raw_data:
        if row[0] != '':
            new_row = [row[0],row[1],row[2],row[3],row[4],row[5]]
            data.append(new_row)
    return data

def factory(data):
    inventory = []
    for row in data:
        splt = SPLT(row)
        inventory.append(splt)
    return inventory

def transplitr(amex_doc, tab):
    data = get_data(amex_doc, tab)
    splits = factory(data)
    for splt in splits:
        print splt.memo


vendor = 'Marriott - San Diego'
transplitr('P061216', '1')