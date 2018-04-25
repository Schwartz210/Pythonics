from spreadsheet import open_file
from os import startfile
from logger import logger
from folder_contents import open_folder

FILE = r'S:\Finance\Job Number Listing\Job Number Listing 041018 ER.xls'
TABS = ['2016','2017','2018']

def get_raw_data():
    '''
    Opens file, gets job info
    '''
    data = []
    for tab in TABS:
        try:
            tab_data = open_file(FILE, tab, 10000)
        except:
            open_folder('job_listing')
            raise Exception('Find correct file name and alter variable FILE on line 6')
        for record in tab_data:
            data.append(record)
    return clean(data)

def clean(raw_data):
    '''
    Removes blank lines
    '''
    data = []
    for record in raw_data:
        if record[3] != '':
            data.append(format_job_code(record))
    return data

def format_job_code(record):
    if type(record[3]) == float:
        record[3] = str(int(record[3]))
    return record

def get_record(job_code, data):
    '''
    Finds the record that matches the desired job_code
    '''
    for record in data:
        if job_code == record[3]:
            return record
    raise Exception(job_code, ': Job not found')

def get_prefixes():
    '''
    Opens Array where prefix data is kept, reads it, returns as dictionary.
    This should be converted to SQL database at some point when I have time
    '''
    prefixes = {}
    file = r'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\Arrays\Client to job prefix.xlsx'
    data = open_file(file, 0, 300)
    del data[0]
    for record in data:
        prefixes[record[0]] = record[1]
    return prefixes

def make_text_string(record):
    '''
    Formats text
    '''
    prefixes = get_prefixes()
    client = record[2]
    job_code = record[3]
    description = record[4]
    try:
        prefix = prefixes[client]
    except:
        startfile(r'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\Arrays\Client to job prefix.xlsx')
        raise Exception('Prefix not found: ', client)
    if '-' not in prefix:
        prefix += '-'
    name = str(prefix + job_code + ' ' + description)
    name = name.replace(' and ',' & ')
    name = name.replace('publication','pub.')
    name = name.replace('Publication','Pub.')
    name = name.replace('meeting','mtg.')
    name = name.replace('Meeting','Mtg.')
    name = name.replace(':',' ')
    name = name.rstrip()
    maximum_char_len_for_quickbooks_name = 41
    return name[:maximum_char_len_for_quickbooks_name]

def create_SQL_statement(full_job, code, client):
    prefixes = get_prefixes()
    prefix = prefixes[client]
    if '-' not in prefix:
        prefix += '-'
    short_job = prefix + code
    SQL_statement = 'INSERT INTO jobs VALUES("%s","%s","","","%s")' % (code, full_job, short_job)
    as_python_function = "execute_sql_command('" + SQL_statement + "')"
    print(full_job)
    print(as_python_function)
    print


@logger
def make_job(job_codes):
    '''
    Container function. This is what gets called in the main module
    '''
    data = get_raw_data()
    for code in job_codes:
        record = get_record(code, data)
        full_job = make_text_string(record)
        create_SQL_statement(full_job, code, record[2])