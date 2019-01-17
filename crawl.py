'''
Written by Abraham "Avi" Schwartz on 01/17/2019 for Python 3.4.
Command line interface for file compression and report emailer daemon.
Usage: python crawl.py directory_arg gmail_address_arg password_arg threshold_int_arg dry_run_bool_arg optional: -h
Kill signal: CTRL + C
'''
from os import listdir
from os.path import join, isfile, getsize
from zipfile import ZipFile, ZIP_DEFLATED
from smtplib import SMTP
from argparse import ArgumentParser
from time import time
from signal import signal, SIGINT

# Initializing global variables
directory = ''
email_address = ''
password = ''
threshold = 0
dry_run = False
files_with_min_compression_gain = []
files_below_threshold = []
files_to_be_compressed = []
total_disc_savings = 0

def to_console(indentations, message):
    '''
    Function for clean console/cmd line output
    '''
    timestamp = time()
    text = ''
    for indention in range(indentations):
        text += '   '
    text += str(timestamp) + ' - ' + str(message)
    print(text)

def signal_handler(signum, frame):
    '''
    SIGTERM kill terminates script with CTRL + C
    '''
    exit()

def setup_argparse():
    '''
    Argparser for entering arguments at the cmd line
    '''
    def data_validation_email(email):
        '''
        Data validation verifies that email address provided is a Gmail address
        '''
        if '@GMAIL.COM' not in email.upper():
            to_console(0, 'This is not a Gmail address: ' + email)
            exit()

    def data_validation_dryrun(dryrun):
        '''
        Data validation for bool passed in a string. Sets global variable dry_run
        '''
        global dry_run
        if dryrun.upper() == 'TRUE':
            dry_run = True
        elif dryrun.upper()  == 'FALSE':
            dry_run = False
        else:
            to_console(0, 'Dry-run not boolean: ' + dryrun)
            exit()

    global directory, email_address, password, threshold, dry_run
    to_console(0, 'Function setup_argparse() is running...')
    parser = ArgumentParser('', add_help=False)
    parser.add_argument('-h', '--help', action='store_true', help='verbose flag')
    parser.add_argument('directory')
    parser.add_argument('email_address')
    parser.add_argument('password')
    parser.add_argument('threshold', type=int)
    parser.add_argument('dry_run_mode')
    args = parser.parse_args()
    if args.help:
        print('~ Help!: Enter args: Directory, Gmail address, Gmail password, Size threshold')
    else:
        print('~ No help')
    directory = args.directory
    email_address = args.email_address
    data_validation_email(email_address)
    password = args.password
    threshold = args.threshold
    data_validation_dryrun(args.dry_run_mode)
    to_console(1, 'Your arguments:')
    to_console(2, 'Directory: ' + directory)
    to_console(2, 'Email addres: ' + email_address)
    to_console(2, 'Password: ' + password)
    to_console(2, 'Threshold: ' + str(threshold))
    to_console(2, 'Dry-run mode: ' + str(dry_run))

def get_files_in_directory():
    '''
    Returns list of files in chosen directory
    '''
    out = []
    to_console(0, 'Function get_files_in_directory() is running...')
    to_console(1, 'Files located it chosen directory:')
    for file_name in listdir(directory):
        if isfile(join(directory, file_name)):
            out.append(file_name)
            to_console(2, file_name)
    return out

def get_file_sizes(files):
    '''
    Returns dictionary of file name and size
    '''
    to_console(0, 'Function get_file_sizes() is running...')
    sizes = {}
    to_console(1, 'Here are the files and their respective sizes:')
    for file_name in files:
        sizes[file_name] = getsize(join(directory, file_name))
        to_console(2, file_name + ' - ' + str(sizes[file_name]))
    return sizes

def compression(file_sizes):
    '''
    Compresses eligible files into zip folders
    '''
    def check_should_compress(file_name):
        '''
        Determines in file should be compressed or left alone. Returns bool
        '''
        global files_with_min_compression_gain, files_below_threshold, files_to_be_compressed
        to_console(0, 'Function check_should_compress() is running...')
        size = file_sizes[file_name]
        to_console(1, file_name + ' is of size ' + str(size))
        if size > threshold:
            to_console(2, file_name + ' exceeds threshold.')
            exceptions = ['Thumbs.db', '.zip', '.jpg']
            for exception in exceptions:
                if exception in file_name:
                    to_console(2, file_name + ' has a min compression gain because it contains: ' + exception)
                    to_console(2, 'Compression required: False')
                    files_with_min_compression_gain.append(file_name)
                    return False
            to_console(2, file_name + ' will be compressed.')
            files_to_be_compressed.append(file_name)
            return True
        else:
            to_console(1, file_name + ' is below the threshold.')
            files_below_threshold.append(file_name)
            return False

    global total_disc_savings
    to_console(0, 'Function compression() is running...')
    for file_name in file_sizes.keys():
        if check_should_compress(file_name):
            zip_file_name = join(directory, file_name) + ' .zip'
            to_console(2, 'New zip folder name: ' + zip_file_name)
            output_zip = ZipFile(zip_file_name, 'w')
            output_zip.write(join(directory, file_name), compress_type=ZIP_DEFLATED)
            output_zip.close()
            original_file_size = file_sizes[file_name]
            compressed_file_size = getsize(zip_file_name)
            to_console(2, 'Computing disk savings...')
            savings = original_file_size - compressed_file_size
            to_console(3, 'Original size - New size = Savings: ' + str(original_file_size) + ' - ' +
                       str(compressed_file_size) + ' = ' + str(savings))
            to_console(3, 'Incrementing total_disc_savings variable with savings for this file.')
            total_disc_savings += savings

def email_report():
    '''
    Sending report via email from Gmail account. If you are having trouble with authentication, then you may need to
    change you Gmail security settings.
    '''
    def get_email_message():
        '''
        Formats text summary for email message
        '''
        to_console(0, 'Function get_email_message() is running...')
        message = '\nThe following files are below the threshold for compression:\n'
        for file_name in files_below_threshold:
            message += '    ' + file_name + '\n'
        message += '\nThe following files are not being compressed due to minimum compression gain:\n'
        for file_name in files_with_min_compression_gain:
            message += '    ' + file_name + '\n'
        message += '\nThe following files are being compressed:\n'
        for file_name in files_to_be_compressed:
            message += '    ' + file_name + '\n'
        message += '\nThe total disk savings was ' + str(total_disc_savings)
        return message

    to_console(0, 'Function email_report() is running...')
    server = SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_address, password)
    message = get_email_message()
    server.sendmail(email_address, email_address, message)
    server.quit()

def execute():
    '''
    Main container function
    '''
    to_console(0, 'Function execute() is running...')
    setup_argparse()
    signal(SIGINT, signal_handler)
    all_files = get_files_in_directory()
    file_sizes = get_file_sizes(all_files)
    if not dry_run:
        compression(file_sizes)
        email_report()
    else:
        to_console(0, 'Dry-run = True. Skipping compression and email...')


if __name__ == '__main__':
    execute()
