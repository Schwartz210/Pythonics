__author__ = 'aschwartz - Schwartz210@gmail.com'
from spreadsheet import open_file, format_date

master_job_dict = {'BDS7' : 'BDSSS77777:',
                   'BDS2' : 'BDSSS22222:'}

def get_subjobs():
    '''
    Opens file, cycles throught tabs pulling subjob data. Returns list of records.
    '''
    file = r"C:\Activ\BDSSSS\BUN\3_Status Report\Status Report_MASTER.xlsx"
    tabs = [2,3,4]
    all_data = []
    for tab in tabs:
        data = open_file(file, tab, 100)
        del data[0]
        for row in data:
            cleansed_row = [row[0], format_date(row[5]),row[9],row[10]]
            all_data.append(cleansed_row)
    return all_data

def get_record(subjobs, job_code):
    '''
    Takes a job code, returns associated record
    '''
    for record in subjobs:
        if record[0] == job_code:
            out = (record[1],record[2],record[3])
            return out
    raise Exception('Job_code not found in data')

def make_job_names(subjobs, job_codes):
    '''
    Returns string job name
    Example: "BDSSSSSSSSS7227 2016 Speaker Bureau Series:389 Naples, FL 08/03/16"
    '''
    names = []
    for job_code in job_codes:
        part1 = master_job_dict[job_code[:7]]
        part2 = job_code[8:11]
        date, city, state = get_record(subjobs, job_code)
        name = part1 + part2 + ' ' + city + ', ' + state + ' ' + date
        names.append(name)
    return names

def subjob_builder(job_codes):
    '''
    Container function. Pulls whole thing together.
    '''
    subjobs = get_subjobs()
    names = make_job_names(subjobs, job_codes)
    for name in names:
        print name


job_codes = ['BDS7227-391']
subjob_builder(job_codes)