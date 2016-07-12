from os import listdir, startfile
from os.path import isfile, join
from shutil import copyfile, move
from error_checking import bold
from window import SingleSelectionWindow
from logger import Logger

financial = 'C:\Files\Financial Files'
plat = financial + r'\AMEX Statements\2016\Platinum'
cent = financial + r'\AMEX Statements\2016\Centurion'

financial_link = r'\\Internal Files\Financial Files'
plat_link = financial_link + r'\AMEX Statements\2016\Platinum'
cent_link = financial_link + r'\AMEX Statements\2016\Centurion'




avi_folder = financial + '\Avi'
root = avi_folder + '\Mass invoice GUI'
folder_dict = {'root' : root,
               'avi' : avi_folder,
               'financial' : financial,
               'amex-02' : r'C:\AMEX Statements\2016\Platinum\02 support docs',
               'amex-03' : r'C:\AMEX Statements\2016\Platinum\03 support docs',
               'amex-04' : r'C:\AMEX Statements\2016\Platinum\04 support docs',
               'amex-05' : r'C:\AMEX Statements\2016\Platinum\05 support docs',
               'amex-06' : r'C:\AMEX Statements\2016\Platinum\06 support docs',
               'payables' : r'C:\Internal Files\Financial Files\Support docs\YR2016\July',
               'h' : 'H:',
               'plat' : plat}

file_dict = {'jobs' : r'H:\User data\Jobs.xlsm',
             'checks' : root + '\Mass check request gui.xlsx',
             'invoices' : root + '\Mass invoice gui.xlsx',
             'amex_prog' : avi_folder + '\Amex Progress.xlsx',
             'allo' : avi_folder + r'\1. Invoice worksheet - Allocation.xlsm',
             'P031316' : plat + '\P031316.xlsm'
             }

amex_dict = {'P011316' : plat + '\P011316',
             'P021116' : plat + '\P021116',
             'P031316' : plat + '\P031316',
             'P041216' : plat + '\P041216',
             'P051316' : plat + '\P051316',
             'P061216' : plat + '\P061216',
             'C012916' : cent + '\C012916',
             'C022616' : cent + '\C022616',
             'C032916' : cent + '\C032916',
             'C042816' : cent + '\C042816',
             'C052916' : cent + '\C052916'
             }
def amex_type(amex_doc):
    if amex_doc[0] == 'P':
        return 'Platinum'
    elif amex_doc[0] == 'C':
        return 'Centurion'
    else:
        return 'ERROR'

def amex_filepath(amex_doc):
    '''
    Example input: 'P061216'
    Example output: 'C:\AMEX Statements\2016\Platinum\P061216'
    '''
    amex = amex_type(amex_doc)
    if amex == 'Platinum': filepath = plat
    elif amex == 'Centurion': filepath = cent
    output =  '%s\%s' % (filepath, amex_doc)
    return output

def amex_support(amex_doc):
    '''
    Input: 'C022616'
    Output: 'Internal Files\Financial Files\01 - AMEX Statements\2016\Centurion\02 support docs'
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
    var = ''
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

    destination = folder_dict[destination_folder]
    print destination
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

def main_file_handler(source_folder, destination_folder):
    '''
    Main subroutine/container function
    '''
    files = get_folder_contents(source_folder)
    files = clean(files)
    copy_files(files, destination_folder)
    print_info(files)

def rename_move_files(source_folder, destination_folder, name):
    files = get_folder_contents(source_folder)
    files = clean(files)
    rename_files(files, name)
    files = get_folder_contents(source_folder)
    files = clean(files)
    copy_files(files, destination_folder)
    print 'The following files are being copied to the support docs folder:\n'
    print_info(files)

def show_folder_contents(source_folder=None):
    if not source_folder: source_folder = input_handler()
    files = get_folder_contents(source_folder)
    files = clean(files)
    print_info(files)
    key_dump(folder_dict)

def open_folder(folder):
    active_folder = folder_dict[folder]
    print '\n\n' + bold('The following folder has been opened:') + '\n' + '%s' % (active_folder)
    startfile(active_folder)
    key_dump(folder_dict)

def open_file2(file):
    active_file = file_dict[file]
    print '\n\n' + bold('The following file has been opened:') + '\n' + '%s' % (active_file)
    startfile(active_file)
    key_dump(file_dict)
    logger_message = 'open_file ran on %s' % (active_file)
    Logger('open_file',logger_message)

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

def open_amex():
    '''
    This functions allows the user to select an Amex, which opens up the PDF and Excel versions.
    '''
    options = amex_dict.keys()
    window = SingleSelectionWindow(options)
    amex = window.selected_options[0]
    worksheet = amex_dict[amex] + '.xlsm'
    statement = amex_dict[amex] + '.pdf'
    startfile(worksheet)
    startfile(statement)

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

