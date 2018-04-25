__author__ = 'aschwartz - Schwartz210@gmail.com'
from spreadsheet import open_file, format_date
from logger import logger

class Profile(object):
    def __init__(self, code, match_job, filepath, tabs, columns, full_job):
        self.code = code
        self.match_job = match_job
        self.filepath = filepath
        self.tabs = tabs
        self.columns = columns
        self.full_job = full_job

    def get_jobs(self):
        all_data = []
        for tab in self.tabs:
            data = open_file(self.filepath, tab, 500)
            del data[0]
            for row in data:
                try:
                    packet = self.build_full_job(row)
                    all_data.append(packet)
                except:
                    pass
        return all_data

    def build_full_job(self, row):
        program_num = row[0][-3:]
        city = row[self.columns[1]]
        state = row[self.columns[2]]
        location = city + ', ' + state
        date = row[self.columns[0]]
        full_job = self.full_job + program_num + ' ' + city + ', ' + state + ' ' + format_date(date)
        code = self.code + '-' + program_num
        short_job = self.match_job + '-' + program_num
        if short_job[:3] == 'BDS':
            short_job = short_job.replace('BDS','BDSI')
        SQL_request = 'INSERT INTO jobs VALUES("%s","%s","%s","%s","%s")' % (code, full_job, format_date(date), location, short_job)
        return [code, full_job, SQL_request]



profiles = [
    Profile('7239',
            'BPL7239',
            r'S:\Active Clients\BPL\7239_Speaker Program Series 2\3_Status Report\BPL7097_7201_7239 - Status and Reg Report_MASTER.xlsx',
            [2,3,4],          #spreadsheet tabs
            [4,8,9],          #spreadsheet columns
            'BPL7239 2016 Speaker Program series 2:'),

    Profile('7270',
            'BDS7270',
            r'S:\Active Clients\BDSI\BUNAVAIL\BDSI 7270 - Speaker Bureau 2017\3_Status Report\BDSI_BUNAVAIL_Speaker Bureau Status Report 2017_MASTER.xlsx',
            [2,3,4],          #spreadsheet tabs
            [5,9,10],         #spreadsheet columns
            'BDSI7270 2017 Bunavail Speaker Bureau:'),

    Profile('7280',
            'BDS7280',
            r'S:\Active Clients\BDSI\BELBUCA\BDSI 7280 - Belbuca Speaker Bureau 2017\3_Status Report\BDSI Belbuca Speaker Bureau Status Report_MASTER.xlsx',
            [2,3,4],          #spreadsheet tabs
            [5,9,10],         #spreadsheet columns
            'BDSI7280 2017 Belbuca Speaker Bureau:'),

    Profile('7340',
            'LEO7340',
            r'S:\Active Clients\Leo\LEO 7340 Live Meeting Series Q3\Status Report\LEO7340_Program Status Report_MASTER.xls',
            [2,3,4],
            [9,16,17],
            'LEO7340 Speaker Bureau Program Q3-Q4:'),

    Profile('8002',
            'BDS8002',
            r'S:\Active Clients\BDSI\BELBUCA\BDSI 8002 - Belbuca Speaker Bureau 2018\3_Status Report\Belbuca Speaker Bureau 2018_Status Report_MASTER.xlsx',
            [2,3,4], #spreadsheet tabs
            [5,9,10], #spreadsheet columns
            'BDSI8002 Belbuca Speaker Bureau 2018:'),

    Profile('7390',
            'BDS7390',
            r'S:\Active Clients\BDSI\BUNAVAIL\BDSI 7390 - Speaker Bureau 2018\3_Status Report\Bunavail Speaker Bureau 2018_Status Report_MASTER.xlsx',
            [2,3,4],
            [5,9,10],
            'BDSI7390 Bunavail 2018 Speaker Bureau Pro:')
    ]

def get_required_profiles(job_numbers):
    out = []
    for number in job_numbers:
        if number[0:4] not in out:
            out.append(number[0:4])
    return out

def get_subjobs(job_numbers):
    '''
    Opens file, cycles through tabs pulling subjob data. Returns list of records.
    '''
    all_data = []
    required_profiles = get_required_profiles(job_numbers)
    for profile in profiles:
        if profile.code in required_profiles:
            data = profile.get_jobs()
            for record in data:
                all_data.append(record)
    return all_data

@logger
def subjob_builder(job_numbers):
    all_jobs = get_subjobs(job_numbers)
    out = []
    for number in job_numbers:
        status = 0
        for job in all_jobs:
            if number == job[0]:
                print(job[2])
                out.append(job[2])
                status = 1
        if status == 0:
            raise Exception('Not found: ', number)
    return out
