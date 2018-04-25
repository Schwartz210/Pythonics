__author__ = 'aschwartz - Schwartz210@gmail.com'
from sqlite3 import connect
from mydatabases import get_client_prefix, get_jobs, get_cc_accts, get_name_class_list, get_general_ledger_list, get_vendors
from openpyxl import Workbook
from tabulate import tabulate
from logger import logger

DATABASE = 'test.db'
TABLES = ['general_ledger','clients', 'jobs','cc_accts','names','vendors']
FIELDS = {'general_ledger':['Short', 'Full', 'Category'],
          'clients':['Short', 'Full'],
          'jobs':['Number', 'Full', 'Date', 'Location', 'Short'],
          'cc_accts':['Short', 'Full'],
          'names':['Full', 'Quickbooks', 'Airline'],
          'vendors':['Name']}

def update_db():
    clients = get_client_prefix()
    jobs = get_jobs()
    cc_accts = get_cc_accts()
    names = get_name_class_list()
    general_ledger = get_general_ledger_list()
    vendors = get_vendors()
    conn = connect(DATABASE)
    c = conn.cursor()
    c.execute('DROP TABLE general_ledger')
    c.execute('DROP TABLE clients')
    c.execute('DROP TABLE jobs')
    c.execute('DROP TABLE cc_accts')
    c.execute('DROP TABLE names')
    c.execute('DROP TABLE vendors')
    c.execute('''CREATE TABLE general_ledger(Short, Full, Category)''')
    c.execute('''CREATE TABLE clients(Short, Full)''')
    c.execute('''CREATE TABLE jobs(Number, Full, Date, Location, Short)''')
    c.execute('''CREATE TABLE cc_accts(Short, Full)''')
    c.execute('''CREATE TABLE names(Full, Quickbooks, Airline)''')
    c.execute('''CREATE TABLE vendors(Name)''')
    c.executemany('INSERT INTO general_ledger VALUES(?,?,?)', general_ledger)
    c.executemany('INSERT INTO clients VALUES(?,?)', clients)
    c.executemany('INSERT INTO jobs VALUES(?,?,?,?,?)', jobs)
    c.executemany('INSERT INTO cc_accts VALUES(?,?)', cc_accts)
    c.executemany('INSERT INTO names VALUES(?,?,?)', names)
    c.executemany('INSERT INTO vendors VALUES(?)', vendors)
    conn.commit()
    conn.close()

def exists(sql_request):
    '''
    Evualuate if record exists. Returns boolean.
    '''
    conn = connect(DATABASE)
    c = conn.cursor()
    count = len(list(c.execute(sql_request)))
    if count > 0:
        out = True
    else:
        out = False
    conn.commit()
    conn.close()
    return out

@logger
def execute_sql_command(SQL_request):
    conn = connect(DATABASE)
    c = conn.cursor()
    try:
        c.execute(SQL_request)
    except:
        print(SQL_request)
    conn.commit()
    conn.close()

def pull_data(SQL_request):
    conn = connect(DATABASE)
    c = conn.cursor()
    try:
        out = list(c.execute(SQL_request))
        conn.commit()
        conn.close()
        return out
    except:
        raise Exception('Not able to fulfill request')

def get_table(table, order):
    qualifier = ''
    if order != '':
        field = FIELDS[table][0]
        order_statement = 'ORDER BY %s ' % (field)
        order_dict = {'a': order_statement + 'ASC', 'd': order_statement + 'DESC'}
        qualifier = order_dict[order]
    SQL_request = 'SELECT * FROM %s %s' % (table, qualifier)
    data = pull_data(SQL_request)
    return data

def get_user_input():
    i = 0
    for table in TABLES:
        print(i, table)
        i += 1
    user_input = raw_input('Pick a table to query.\n')
    acceptable_inputs = [str(j) for j in range(len(TABLES))]
    while user_input not in acceptable_inputs:
        user_input = raw_input('Pick a table to query.\n')
    return int(user_input)

def get_table_name():
    '''
    Collects table name selection from user
    '''
    user_input = get_user_input()
    table = TABLES[user_input]
    return table

def get_modifier():
    acceptable_orders = ['a','d','']
    order_input = raw_input('Order ascending(a) or descending(d) or enter nothing to skip: ')
    while order_input not in acceptable_orders:
        order_input = raw_input('Order ascending("a") or descending("d") or enter nothing to skip: ')
    return order_input

@logger
def query(table_name=None):
    if not table_name:
        table_index = get_user_input()
        table_name = TABLES[table_index]
    modifier = get_modifier()
    table = get_table(table_name, modifier)
    headers = FIELDS[table_name]
    print tabulate(table, headers=headers)

def insert_into(SQL_request=None):
    if SQL_request:
        execute_sql_command(SQL_request)
        update_spreadsheets()
        return
    user_input = get_user_input()
    query(TABLES[user_input])
    table_name = TABLES[user_input]
    fields = FIELDS[table_name]
    values = ''
    print 'Type "exit" to exit'
    for field in fields:
        user_value = raw_input(field + ':\n')
        if user_value == 'exit': return
        values += "'" + user_value + "',"
    SQL_request = "INSERT INTO %s VALUES(%s)" % (table_name, values[:len(values)-1])
    print SQL_request
    execute_sql_command(SQL_request)
    update_spreadsheets()

def convert_to_int_if_applicable(value):
    try:
        return int(value)
    except:
        return value

def convert_tuples_to_list(list_of_tuples):
    new_list = []
    for tup in list_of_tuples:
        lower_list = []
        for value in tup:
            lower_list.append(value)
        new_list.append(lower_list)
    return new_list

def update_workbook(filepath, values):
    wb = Workbook()
    ws = wb.active
    for value in values:
        ws.append(value)
    wb.save(filepath)

def update_spreadsheets():
    for table in TABLES:
        SQL_request = 'SELECT * FROM %s' % (table)
        data = pull_data(SQL_request)
        if table == 'jobs':
            data = convert_tuples_to_list(data)
            for record in data:
                record[0] = convert_to_int_if_applicable(record[0])
        if table == 'names':
            data = convert_tuples_to_list(data)
            for record in data:
                record.append(record[0])
                record.append(record[1])
        name = r"S:\CRG Internal Files\Financial Files\Avi\Exports\%s.xlsx" % (table)
        update_workbook(name, data)
    return

def commit_subjob_to_database(SQL_command):
    acceptable_answers = ['y','n']
    print SQL_command
    answer = raw_input('Do you want to commit this to the database (y or n)?\n')
    while answer not in acceptable_answers:
        answer = raw_input('Do you want to commit this to the database (y or n)?\n')
    if answer == 'y':
        execute_sql_command(SQL_command)

def handler_subjobs(subjobs):
    for job in subjobs:
        commit_subjob_to_database(job)
    update_spreadsheets()
    return subjobs

def print_tables():
    con = connect(DATABASE)
    cursor = con.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    print(cursor.fetchall())

@logger
def update_job(number, **kwargs):
    if 'Date' in kwargs:
        SQL_request = 'UPDATE Jobs SET '
        SQL_request += 'Date="' + kwargs['Date'] + '" '
        SQL_request += 'WHERE Number="' + number + '"'
        execute_sql_command(SQL_request)
    if 'Location' in kwargs:
        SQL_request = 'UPDATE Jobs SET '
        SQL_request += 'Location="' + kwargs['Location'] + '" '
        SQL_request += 'WHERE Number="' + number + '"'
        execute_sql_command(SQL_request)
    if 'Full' in kwargs:
        SQL_request = 'UPDATE Jobs SET '
        SQL_request += 'Full="' + kwargs['Full'] + '" '
        SQL_request += 'WHERE Number="' + number + '"'
        execute_sql_command(SQL_request)

def insert_many_names(records):
    conn = connect(DATABASE)
    c = conn.cursor()
    for record in records:
        SQL_request = 'INSERT INTO names VALUES("%s", "%s", "%s")' % (record[0], record[1], record[2])
        c.execute(SQL_request)
    conn.commit()
    conn.close()

@logger
def update_name(full, **kwargs):
    if 'Quickbooks' in kwargs:
        SQL_request = 'UPDATE Names SET '
        SQL_request += 'Quickbooks="' + kwargs['Quickbooks'] + '" '
        SQL_request += 'WHERE Full="' + full + '"'
        execute_sql_command(SQL_request)
    if 'Airline' in kwargs:
        SQL_request = 'UPDATE Names SET '
        SQL_request += 'Airline="' + kwargs['Airline'] + '" '
        SQL_request += 'WHERE Full="' + full + '"'
        execute_sql_command(SQL_request)

def get_location(full_job):
    SQL_request = 'SELECT Location FROM Jobs WHERE Full="%s"' % (full_job)
    try:
        data = pull_data(SQL_request)[0][0]
    except:
        raise Exception('Full job not founds: ', full_job)
    return data
