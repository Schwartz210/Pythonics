from spreadsheet import open_file
from super_transaction import SuperTransaction
from xlwt import Workbook
from collections import Counter


class HotelLineItem(SuperTransaction):
    def __init__(self, packet, vendor):
        self.gl_in = packet[0]
        self.person = packet[1]
        self.amount = packet[2]
        self.date = self.format_date(packet[3])
        self.comments = packet[4]
        self.job_in = self.str_format(packet[5])
        self.vendor = vendor
        self.full_job = self.find_full_job_name(self.job_in)
        self.job_out = self.create_job_out(self.full_job, self.gl_in)
        self.client = self.find_client(self.full_job)
        self.gl_out = self.create_gl_out(self.gl_in, self.client)
        self.memo = self.create_memo()
        self.person_class = self.find_person_class(self.person)

    def __str__(self):
        print self.gl_out, self.amount, self.memo, self.job_out, self.person_class

    def create_memo(self):
        '''
        String formatting for memo field.
        '''
        memo = self.vendor + ' - ' + self.gl_in
        optional_fields = [self.person, self.comments, self.date]
        for field in optional_fields:
            if field:
                memo += ' - ' + field
        return memo

    def to_xl(self,sheet, row):
        sheet.write(row, 0,self.gl_out)
        sheet.write(row, 1,self.amount)
        sheet.write(row, 2,self.memo)
        if self.job_out:
            sheet.write(row, 3,self.job_out)
        sheet.write(row, 4,self.person_class)

class HotelFactory(object):
    def __init__(self, filename, vendor, invoice_date, job_in, program_manager):
        self.filename = filename
        self.vendor = vendor
        self.invoice_date = invoice_date
        self.job_in = job_in
        self.program_manager = program_manager
        self.data = self.get_data(self.filename)
        self.line_items = self.obj_factory(self.data)
        self.total_amount = self.get_total()
        self.publish(self.line_items)
        self.find_missing_person_class()

    def get_data(self, filename):
        '''
        This method pulls the data from the spreadsheet
        '''
        data = open_file(filename,'py',500)
        del data[0]
        return data

    def obj_factory(self, data):
        '''
        This method is the factory. It creates instances of HotelLineItem class.
        '''
        output_list = []
        for row in data:
            if row[0]:
                line_item = HotelLineItem(row, self.vendor)
                output_list.append(line_item)
        return output_list

    def get_total(self):
        '''
        Returns sum of line item amounts
        '''
        total = 0
        for line_item in self.line_items:
            total += line_item.amount
        return total

    def publish(self, list):
        '''
        This method creates the output spreadsheet.
        '''
        extra_title = '$' + str(self.total_amount)
        book = Workbook()
        sh = book.add_sheet('QB')
        sh.write(0, 0,'GL')
        sh.write(0, 1,'Amount')
        sh.write(0, 2,'Memo')
        sh.write(0, 3,'Job')
        sh.write(0, 4,'Class')
        i = 1
        for obj in list:
            obj.to_xl(sh, i)
            i += 1
        name = self.vendor + ' - ' + extra_title + '.xls'
        book.save('H:\%s' % (name))

    def find_missing_person_class(self):
        '''
        This method collects all people with unknown QB class data and prints to terminal, as a form of error check.
        '''
        unique = set()
        for line_item in self.line_items:
            if line_item.person_class == 'ERROR':
                unique.add(line_item.person)

        if len(unique) > 0:
            print
            print 'These people need to be added to the Name-Class Array'
            for person in unique:
                print person


class CoversheetFactory(object):
    '''
    This class is used to create the coversheets for hotel bills.
    '''
    def __init__(self, HotelFactory):
        self.line_items = HotelFactory.line_items
        self.job_in = HotelFactory.job_in
        self.program_manager = HotelFactory.program_manager
        self.cover_sheet_rows = []
        self.person_line_items = []
        self.banquet_line_items = []
        self.deposit_line_items = []
        self.remaining_line_items = []
        self.coversheet_row = 4
        self.buckets()
        self.collect_rows()
        self.publish()

    def get_room_count_mode(self):
        '''
        Returns the number of rooms and the mode lodging rate.
        '''
        room_count = 0
        room_amounts = []
        for line_item in self.line_items:
            if line_item.gl_in == 'Lodging' and line_item.person != 'Deposit':
                room_count += 1
                room_amounts.append(line_item.amount)
        counter = Counter(room_amounts)
        mode = counter.most_common(1)[0][0]
        return (room_count, mode)

    def create_headers(self):
        '''
        Create the first three rows of the spreadsheet, which are just headers.
        '''
        output_list = []
        room_count, mode = self.get_room_count_mode()
        local_supertran = SuperTransaction()
        first_row = [vendor, None, None, None, None, invoice_date, 'Job#', local_supertran.find_short_job(self.job_in), None, None, 'Program Man.',self.program_manager]
        second_row = [None, '# of Rooms',room_count, None, 'Rate per room', mode, None, None, '(Room + Tax)',None, None,None]
        third_row = ['Guest','Audio Visual','Faculty Amenities','Banquet','Misc. On-Site','Ground', 'Telephone','Per Diem','Room Chgs','Room Rental','CRG Expense','Total']
        output_list.append(first_row)
        output_list.append(second_row)
        output_list.append(third_row)
        return output_list

    def buckets(self):
        '''
        This methods places every line item into a bucket which determines it's handling.
        '''
        for line_item in self.line_items:
            if line_item.person == 'Deposit':
                self.deposit_line_items.append(line_item)
            elif line_item.gl_in == 'Lodging' or line_item.gl_in == 'Per Diem' or line_item.gl_in == 'Ground':
                self.person_line_items.append(line_item)
            elif line_item.comments != '':
                self.banquet_line_items.append(line_item)
            else:
                self.remaining_line_items.append(line_item)

    def collect_rows(self):
        '''
        The method calls all the row handlers and combines the results into a list of rows. It updates the class-
        variable 'self.cover_sheet_rows'.
        '''
        header_rows = self.create_headers()
        person_rows = self.person_line_handler()
        banquet_rows = self.banquet_line_handler()
        remaining_rows = self.remaining_line_handler()
        bottom_totals = self.column_totals()
        deposit_row = self.deposit_line()
        balance_row = self.balance_row()
        collection = [header_rows, person_rows, banquet_rows,remaining_rows,bottom_totals,deposit_row,balance_row]
        for packet in collection:
            for row in packet:
                self.cover_sheet_rows.append(row)


    def person_line_handler(self):
        '''
        This row handler handles Lodging, Per Diem, and Ground charges. It create a single row for each person. and-
        aggregates those costs into that row.
        '''
        output_list = []
        names = []
        for line_item in self.person_line_items:
            if line_item.person not in names:
                excel_formula = '=sum(B%s:K%s)' % (self.coversheet_row, self.coversheet_row)
                self.coversheet_row += 1
                names.append(line_item.person)
                if line_item.gl_in == 'Ground':
                    row = [line_item.person,None,None,None,None,line_item.amount,None,0.0,0.0, None,None, excel_formula]
                elif line_item.gl_in == 'Per Diem':
                    row = [line_item.person,None,None,None,None,0.0,None,line_item.amount,0.0, None,None, excel_formula]
                elif line_item.gl_in == 'Lodging':
                    row = [line_item.person,None,None,None,None,0.0,None,0.0,line_item.amount, None,None, excel_formula]
                else:
                    row = ['ERROR']
                output_list.append(row)
            else:
                for row in output_list:
                    if line_item.person == row[0]:
                        if line_item.gl_in == 'Ground': row[5] += line_item.amount
                        elif line_item.gl_in == 'Per Diem': row[7] += line_item.amount
                        elif line_item.gl_in == 'Lodging': row[8] += line_item.amount
        return output_list

    def banquet_line_handler(self):
        '''
        This handler handles banquet charges. This includes F&B, and potentially A/V, Room Rental, and Misc.
        '''
        output_list = []
        names = []
        for line_item in self.banquet_line_items:
            if line_item.comments not in names:
                excel_formula = '=sum(B%s:K%s)' % (self.coversheet_row, self.coversheet_row)
                self.coversheet_row += 1
                names.append(line_item.comments)
                if line_item.gl_in == 'AudioVisual':
                    row = [line_item.comments,line_item.amount,None,None,None,None,None,None,None,None,None,excel_formula]
                elif line_item.gl_in == 'Food & Beverage':
                    row = [line_item.comments,None,None,line_item.amount,None,None,None,None,None,None,None,excel_formula]
                elif line_item.gl_in == 'Misc':
                    row = [line_item.comments,None,None,None,line_item.amount,None,None,None,None,None,None,excel_formula]
                elif line_item.gl_in == 'Room Rental':
                    row = [line_item.comments,None,None,None,None,None,None,None,None,line_item.amount,None,excel_formula]
                else:
                    row = ['ERROR']
                output_list.append(row)
            else:
                for row in output_list:
                    if line_item.comments == row[0]:
                        if line_item.gl_in == 'AudioVisual': row[1] += line_item.amount
                        elif line_item.gl_in == 'Food & Beverage': row[3] += line_item.amount
                        elif line_item.gl_in == 'Misc': row[4] += line_item.amount
                        elif line_item.gl_in == 'Room Rental': row[5] += line_item.amount
        return output_list

    def remaining_line_handler(self):
        '''
        This handler handles everything that 1) isn't a person-line-item or 2) isn't part of a banquet invoices.
        '''
        output_list = []
        for line_item in self.remaining_line_items:
            excel_formula = '=sum(B%s:K%s)' % (self.coversheet_row, self.coversheet_row)
            self.coversheet_row += 1
            row = [line_item.gl_in,0.0,None,None,0.0,None,0.0,None,None,0.0,None, excel_formula]
            if line_item.gl_in == 'AudioVisual': row[1] = line_item.amount
            elif line_item.gl_in == 'Misc': row[4] = line_item.amount
            elif line_item.gl_in == 'Telephone': row[6] = line_item.amount
            elif line_item.gl_in == 'Room Rental': row[9] = line_item.amount
            else: print 'ERROR in remaining_line_handler()'
            output_list.append(row)
        return output_list

    def column_totals(self):
        '''
        This method produces a spreadsheet row that has the totals (Excel formula) for the columns.
        '''
        col_a = 'TOTAL'
        col_b = '=sum(B4:B%s)' % (self.coversheet_row-1)
        col_c = '=sum(C4:B%s)' % (self.coversheet_row-1)
        col_d = '=sum(D4:B%s)' % (self.coversheet_row-1)
        col_e = '=sum(E4:B%s)' % (self.coversheet_row-1)
        col_f = '=sum(F4:B%s)' % (self.coversheet_row-1)
        col_g = '=sum(G4:B%s)' % (self.coversheet_row-1)
        col_h = '=sum(H4:B%s)' % (self.coversheet_row-1)
        col_i = '=sum(I4:B%s)' % (self.coversheet_row-1)
        col_j = '=sum(J4:B%s)' % (self.coversheet_row-1)
        col_k = '=sum(K4:B%s)' % (self.coversheet_row-1)
        col_l = '=sum(L4:B%s)' % (self.coversheet_row-1)
        row = [col_a,col_b,col_c,col_d,col_e,col_f,col_g,col_h,col_i,col_j,col_k,col_l]
        return [row]

    def deposit_line(self):
        '''
        Produces a row for the deposits
        '''
        total = 0.0
        for line_item in self.deposit_line_items:
            total -= line_item.amount
        row = [None,None,None,None,None,None,None,None,None,None,'Credits/Deposit',total]
        return [row]

    def balance_row(self):
        '''
        Produces a row on the spreadsheet which contains the balance (Amount to be paid).
        Grand Total - Deposits = Balance
        '''
        excel_formula = '=L%s-L%s' % (self.coversheet_row,self.coversheet_row+1)
        row = [None,None,None,None,None,None,None,None,None,None,'Balance',excel_formula]
        return [row]

    def publish(self):
        '''
        This method creates the output spreadsheet.
        '''
        book = Workbook()
        sh = book.add_sheet('coversheet')
        r = 0
        c = 0
        for row in self.cover_sheet_rows:
            for item in row:
                sh.write(r, c,item)
                c += 1
            r += 1
            c = 0
        name = 'coversheet.xls'
        book.save('H:\%s' % (name))

file = r"C:\Internal Files\Financial Files\Hotel Bill\worksheet.xlsx"
vendor = 'Hilton Columbus'
invoice_date = '04.26.16'
program_manager = 'Bob'
job_in = 7211
h = HotelFactory(file, vendor, invoice_date, job_in, program_manager)
#coversheet = CoversheetFactory(h)