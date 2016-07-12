from folder_contents import financial
from spreadsheet import open_file, keep_these_columns

database_directory = financial + r"\01 - AMEX Statements\Arrays"

def get_client_prefix():
    filepath = database_directory + '\Client.xlsx'
    data = open_file(filepath,0,150)
    return data

def get_jobs():
    filepath = 'H:\User data\Jobs.xlsm'
    data = open_file(filepath,0,3500)
    return data

def get_cc_accts():
    filepath = database_directory + '\CC Accts.xlsx'
    data = open_file(filepath,0,50)
    return data

def get_name_class_list():
    filepath = database_directory + '\Name-class array.xlsx'
    data = open_file(filepath,0,3000)
    return data

def get_general_ledger_list():
    filepath = database_directory + '\GL-Short descrip array.xlsx'
    data = open_file(filepath,1,230)
    return data

def get_address_list():
    filepath = financial + r'\Avi\Address.xlsx'
    data = open_file(filepath,1,15000)
    data = keep_these_columns(data,0,2,4,6,8,10,12,14,16,18,28)
    return data


clients = get_client_prefix()
jobs = get_jobs()
cc_accts = get_cc_accts()
names = get_name_class_list()
general_ledger = get_general_ledger_list()
addresses = get_address_list()
