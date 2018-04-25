from transaction_pull import get_transactions
from spreadsheet import col_to_num
from logger import logger

class Tran(object):
    def __init__(self, record):
        self.trans_number = record[0]
        self.memo = str(record[2]).replace(',','')
        self.amount = record[4]
        self.client = record[3] #account
        self.bucket = self.set_bucket(record[1]) #job
        self.num = record[5]

    def set_bucket(self, job):
        buckets = {'Fed Ex' : 'Fed Ex',
                   'Ground' : 'Ground',
                   'Per Diem' : 'Per Diem',
                   'Telephone' : 'Telephone',
                   'Airfare' : 'Airfare',
                   'Food & Beverage' : 'Food & Beverage',
                   'Food & Bevarage' : 'Food & Beverage',
                   'Room Rental' : 'Room Rental',
                   'AudioVisual' : 'AudioVisual',
                   'Audio Visual' : 'AudioVisual',
                   'TrvlMgmtFee' : 'TrvlMgmtFee',
                   'TrvlMgmntFee' :'TrvlMgmtFee',
                   'TrvlMngmntFee' :'TrvlMgmtFee',
                   'Fee For Service' : 'Fee For Service',
                   'Fee for Service' : 'Fee For Service',
                   'Fee FOr Service' : 'Fee For Service',
                   'Lodging' : 'Lodging',
                   'Fees' : 'Fees',
                   'References' : 'References',
                   'Creative Services' : 'Creative Services',
                   'Postage' : 'Postage',
                   'Administrative Services' : 'Administrative Services',
                   'Admiistrative Services' : 'Administrative Services',
                   'Adminstrative Services' : 'Administrative Services',
                   'Administrative Support' : 'Administrative Services',
                   'Administative Services' : 'Administrative Services',
                   'Administrative Service' : 'Administrative Services',
                   'Administrative Sevices' : 'Administrative Services',
                   'Adminitrative Services' : 'Administrative Services',
                   'Administrtaive Services' : 'Administrative Services',
                   'Administrattive Services' : 'Administrative Services',
                   'Administration' : 'Administrative Services',
                   'Administrative Servces' : 'Administrative Services',
                   'Audience Recruitment' : 'Audience Recruitment',
                   'Commission': 'Commission',
                   'Commisson' : 'Commission',
                   'commission': 'Commission',
                   'Creative' : 'Creative Services',
                   'Editorial' : 'Editorial',
                   'Meeting Planning' : 'Meeting Planning',
                   'Misc' : 'Misc',
                   'Site Inspection' : 'Site Inspection',
                   'Transcription':'Transcription',
                   'Production' : 'Production',
                   'Translation' : 'Translation',
                   'Registration' : 'Registration',
                   'Labor':'Labor',
                   'PPN7245 PPN Patient Educ Initiative:Marketing':'PPN',
                   'Telephine' : 'Telephone',
                   'PPN7245 PPN Patient Educ Initiative:Professional Services':'PPN'}
        for key in buckets.keys():
            if key in job:
                return buckets[key]
        raise Exception('No bucket found: ', job, self.trans_number, self.memo)

    def __str__(self):
        return str(self.amount) + ',' + self.client + ',' + self.bucket + ',' + str(self.trans_number) + ',' + self.memo + ',' + self.num + ',' + '\n'


def factory(records):
    out = []
    for record in records:
        out.append(Tran(record))
    return out

def publish(records):
    csv = open('H:\output.csv', 'w')
    for record in records:
        try:
            csv.write(record.__str__())
        except:
            raise Exception('ascii error', record.memo)
    csv.close()

@logger
def meeting_cost_execute():
    raw_data = get_transactions(r'H:\Book4.xlsx', mandatory_column=col_to_num('H'), keep_columns=[col_to_num('H'), #trans num
                                                                                                  col_to_num('N'), #job
                                                                                                  col_to_num('P'), #memo
                                                                                                  col_to_num('R'), #client
                                                                                                  col_to_num('T'),
                                                                                                  col_to_num('L')]) #amount
    transactions = factory(raw_data)
    publish(transactions)


meeting_cost_execute()
