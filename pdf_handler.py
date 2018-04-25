import PyPDF2
from spreadsheet import open_file
from folder_contents import amex_filepath
from logger import logger

class PDFHandler(object):
    '''
    This class takes a Amex PDF and produces a collections of PDF's with particular pages. Each of the new PDF's is-
    meant to be distributed to different staff members based on who card the charges belong to.
    '''
    def __init__(self, amex):
        self.amex = amex
        self.raw_data = self.amex_refnum_pull()
        self.accts = self.get_unique_accts()
        self.refnum_acct_combos = self.get_refnum_acct_combos()
        self.accts_to_pages = self.create_accts_to_pages_dict()
        self.pdf_container()

    def amex_refnum_pull(self):
        '''
        This function returns a list of refnums and accounts from the Amex Excel workbook.
        '''
        filepath = amex_filepath(self.amex) + '.xlsm'
        file = open_file(filepath,0,1000)
        new_list = []
        for row in file[1:]:
            new_record = [row[0],row[3]]
            new_list.append(new_record)
        return new_list

    def get_refnum_acct_combos(self):
        '''
        Returns a list of unique combination of Accounts and Refnums. Used to associate a CRG associate with the-
        amex pages that they should receive to code.
        '''
        unique = []
        for row in self.raw_data:
            if row[0] != 'Main':
                if row not in unique:
                    unique.append(row)
        return unique

    def get_unique_accts(self):
        '''
        Returns list of unique accts
        '''
        accts = set()
        for row in self.raw_data:
            accts.add(row[0])
        return list(accts)

    def create_accts_to_pages_dict(self):
        '''
        Creates a dictionary associating accts with page numbers
        '''
        accts_to_pages = {}
        for acct in self.accts:
            local_list = []
            for row in self.refnum_acct_combos:
                if row[0] == acct:
                    local_list.append(row[1][-2:])
            accts_to_pages[acct] = local_list
        return accts_to_pages

    def create_reader(self):
        filepath = amex_filepath(self.amex) + '.pdf'
        try:
            pdfFileObj = open(filepath, 'rb')
        except:
            raise Exception('Open Foxit, then print to PDF using "PDFCreator". Make sure the file is saved in the correct directory.')
        return pdfFileObj

    def writer(self, pdfReader, key, dict):
        pdfWriter = PyPDF2.PdfFileWriter()
        pdfFileObj = open('H:\AMEX request.pdf', 'rb')
        coversheet_file = PyPDF2.PdfFileReader(pdfFileObj)
        coversheet = coversheet_file.getPage(0)
        pdfWriter.addPage(coversheet)
        pages = dict[key]
        for page in pages:
            if '.' in page:
                raise Exception('One of the reference numbers does not correspond to the expected pattern. Ref numbers '
                                'are supposed to have two digits passed the decimal, this one has fewer.')
            try:
                pageObj = pdfReader.getPage(int(page)-1)
            except:
                raise Exception('The file is fucked up, not the code. Open Foxit, then print to PDF using "PDFCreator".')
            pdfWriter.addPage(pageObj)
        pdfOutputFile = open(r'H:\f\amex ' + key + '.pdf', 'wb')
        print key
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()

    def pdf_container(self):
        pdfFileObj = self.create_reader()
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        for key in self.accts_to_pages.keys():
            self.writer(pdfReader,key, self.accts_to_pages)
        pdfFileObj.close()



@logger
def run_pdf_handler(amex):
    PDFHandler(amex)