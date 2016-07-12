import smtplib
import win32com.client
from folder_contents import get_folder_contents, send_files_to_printer
from logger import Logger
from os import startfile

def send_email():
    sender = 'email@gmail.com'
    receivers = ['email@gmail.com']
    message = """From: Avi Schwartz <email@gmail.com>
    To: To Person <to@todomain.com>
    Subject: SMTP e-mail test

    This is a test e-mail message.
    """

    try:
        smtpObj = smtplib.SMTP('email@gmail.com')
        smtpObj.sendmail(sender, receivers, message)
        print "Successfully sent email"
    except:
        print "Error: unable to send email"

def retrieve_last_ten_emails_from_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    deleted_items = outlook.GetDefaultFolder(3)
    sent_items = outlook.GetDefaultFolder(5)
    messages = inbox.Items
    for i in range(25):
        counter = len(messages) - 1 - i
        message = messages[counter]
        print message.Subject

def is_check_request(message):
    '''
    This function evualates to an email's subject line and returns True if it contains a 'magic_word', else False.
    '''
    magic_words = ['Chk','Speaker expense','Check request','Check Request','FFS Request','check request','BDSI Requests',
                   'BDSI Expense','Expense Request','BDSI Reqs','Expense Form', 'BDSI FFS Chk Req', 'Expense Check']
    for word in magic_words:
        if message.Subject.find(word) >= 0:
            return True
    return False

def emails():
    '''
    Goes through unread emails looking for check requests. Downloads and prints check requests
    '''
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    for i in range(len(messages)-1):
        counter = len(messages) - 1 - i
        message = messages[counter]
        if message.Unread == True:
            if is_check_request(message):
                logger_msg = message.SenderName + " - " + message.Subject + " - " + str(message.LastModificationTime) + '\n'
                Logger('emails', logger_msg)
                if message.Attachments.Count>0:
                    attachments = message.Attachments #assign attachments to attachment variable
                    for attachment in attachments:
                        doc_name = str(attachment)
                        if doc_name.find('.jpg') < 0:
                            attachment.SaveASFile("H:/%s" % (attachment))
                            startfile(attachment, "print")
                            print 'File Saved : ', attachment
                    message.Categories = 'Purple Category'
                    message.Unread = False


