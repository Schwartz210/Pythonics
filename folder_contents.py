from os import listdir, startfile
from os.path import isfile, join
from shutil import copyfile, move
from error_checking import bold
from logger import logger

financial = 'S:\CRG Internal Files\Financial Files'
plat = financial + r'\01 - AMEX Statements\2018\Platinum'
cent = financial + r'\01 - AMEX Statements\2018\Centurion'

financial_link = r'\\crgnas01\shared\CRG Internal Files\Financial Files'
plat_link = financial_link + r'\01 - AMEX Statements\2018\Platinum'
cent_link = financial_link + r'\01 - AMEX Statements\2018\Centurion'

avi_folder = financial + '\Avi'
root = avi_folder + '\Pythonics'
folder_dict = {'root' : root,
               'avi' : avi_folder,
               'financial' : financial,
               'payables' : r'S:\CRG Internal Files\Financial Files\Support docs\YR2018\April',
               'h' : 'H:',
               'plat' : plat,
               'cent' : cent,
               'job_listing' : r'S:\Finance\Job Number Listing'}

file_dict = {'jobs' : r'H:\User data\Jobs.xlsm',
             'checks' : root + '\Mass check request gui.xlsx',
             'invoices' : root + '\Mass invoice gui.xlsx',
             'amex_prog' : avi_folder + '\Amex Progress.xlsx',
             'allo' : avi_folder + r'\1. Invoice worksheet - Allocation.xlsm',
             'vendor': r'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\Arrays\Vendor name array.xlsm',
             'airport' : r'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\Arrays\Location array.xlsx',
             'support' : r'S:\CRG Internal Files\Financial Files\Support docs'}

def amex_type(amex_doc):
    if amex_doc[0] == 'P':
        return 'Platinum'
    elif amex_doc[0] == 'C':
        return 'Centurion'
    else:
        return 'ERROR'

def is_amex(amex_doc):
    '''
    Evualates if argument is an Amex reference of not.
    '''
    if amex_doc[0] != 'C' and amex_doc[0] != 'P':
        return False
    elif not len(amex_doc) == 7:
        return False
    elif not 0 < int(amex_doc[1:3]) < 13:
        return False
    elif not 14 < int(amex_doc[5:7]) < 19:
        return False
    else:
        return True

def amex_filepath(amex_doc):
    '''
    Example input: 'P061216'
    Example output: 'S:\CRG Internal Files\Financial Files\01 - AMEX Statements\2016\Platinum\P061216'
    '''
    amex = amex_type(amex_doc)
    if amex == 'Platinum': filepath = plat
    elif amex == 'Centurion': filepath = cent
    output =  '%s\%s' % (filepath, amex_doc)
    return output

def amex_support(amex_doc):
    '''
    Input: 'C022616'
    Output: '\\crgnas01\shared\CRG Internal Files\Financial Files\01 - AMEX Statements\2016\Centurion\02 support docs'
    '''
    amex = amex_type(amex_doc)
    my_dict = {'Platinum' : plat_link, 'Centurion' : cent_link}
    month = amex_doc[1:3]
    output = '%s\%s support docs' % (my_dict[amex], month)
    return output


def get_folder_contents(folder):
    '''
    Creates a list of all files in the H Drive
    '''
    mypath = folder_dict[folder]
    file_list = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    return file_list

def clean(file_list):
    '''
    Removes unwanted files from list
    '''
    output_list = []
    for file in file_list:
        if file != 'Thumbs.db' and file[-4:] != '.iif' and file[0] != '~':
            output_list.append(file)
    return output_list

def copy_files(file_list, destination_folder):
    '''
    Copies files from H Drive to support doc folder
    '''
    if is_amex(destination_folder):
        destination = amex_support(destination_folder)
    else:
        destination = folder_dict[destination_folder]
    for file in file_list:
        src = 'H:\%s' % (file)
        file_dest = destination + '\%s' % (file)
        copyfile(src, file_dest)

def rename_files(file_list, name):
    series = 1
    for file in file_list:
        if file[:5].lower() == 'sharp':
            src = 'H:\%s' % (file)
            new_file = name + str(series) + '.pdf'
            dest = 'H:\%s' % new_file
            series += 1
            move(src, dest)

def print_info(file_list):
    for file in file_list:
        print file

@logger
def main_file_handler(source_folder, destination_folder):
    '''
    Main subroutine/container function
    '''
    files = get_folder_contents(source_folder)
    files = clean(files)
    copy_files(files, destination_folder)
    print_info(files)

@logger
def rename_move_files(source_folder, destination_folder, name):
    files = get_folder_contents(source_folder)
    files = clean(files)
    rename_files(files, name)
    files = get_folder_contents(source_folder)
    files = clean(files)
    copy_files(files, destination_folder)
    print 'The following files are being copied to the support docs folder:\n'
    print_info(files)

@logger
def show_folder_contents(source_folder=None):
    if not source_folder: source_folder = input_handler()
    files = get_folder_contents(source_folder)
    files = clean(files)
    print_info(files)
    key_dump(folder_dict)

@logger
def open_folder(folder):
    active_folder = folder_dict[folder]
    print '\n\n' + bold('The following folder has been opened:') + '\n' + '%s' % (active_folder)
    startfile(active_folder)
    key_dump(folder_dict)

@logger
def open_file2(file):
    active_file = file_dict[file]
    print '\n\n' + bold('The following file has been opened:') + '\n' + '%s' % (active_file)
    startfile(active_file)
    key_dump(file_dict)

def key_dump(dictionary):
    print bold('\n\nAvailable options are as follows:')
    for key in dictionary.keys():
        print '    ' + key

def input_handler():
    print 'Type in source folder name: \n'
    print 'Your options are: \n'
    for key in folder_dict.keys():
        print '    ' + key
    user_input = raw_input('Answer: \n')
    if user_input not in folder_dict.keys():
        print '%s is not an acceptable answer.' % (user_input)
        return input_handler()
    else:
        return user_input

def send_files_to_printer(files):
    '''
    Sends files to default printer. Creates paper copies.
    '''
    files = clean(files)
    print 'The following files have been sent to the printer. Please allow three minutes to print:','\n'
    for file in files:
        full_file = r'H:' + '\%s' % (file)
        try:
            startfile(full_file, "print")
        except:
            print 'ERROR: %s did not print' % (file)




