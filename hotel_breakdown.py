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
    def __init__(self, HotelFactory):
        self.line_items = HotelFactory.line_items
        self.remaining_line_items = list(self.line_items)
        self.job_in = HotelFactory.job_in
        self.program_manager = HotelFactory.program_manager
        self.cover_sheet_rows = []
        self.coversheet_row = 4
        self.create_cover_sheet()

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

    def create_cover_sheet(self):
        room_count, mode = self.get_room_count_mode()
        local_supertran = SuperTransaction()
        first_row = [vendor, None, None, None, None, invoice_date, 'Job#', local_supertran.find_short_job(self.job_in), None, None, 'Program Man.',self.program_manager]
        second_row = [None, '# of Rooms',room_count, None, 'Rate per room', mode, None, None, '(Room + Tax)',None, None,None]
        third_row = ['Guest','Audio Visual','Faculty Amenities','Banquet','Misc. On-Site','Ground', 'Telephone','Per Diem','Room Chgs','Room Rental','CRG Expense','Total']
        self.cover_sheet_rows.append(first_row)
        self.cover_sheet_rows.append(second_row)
        self.cover_sheet_rows.append(third_row)
        person_rows = self.coversheet_person_rows()
        for row in person_rows:
            self.cover_sheet_rows.append(row)
        self.coversheet_banquet()

    def coversheet_per_person(self):
        '''
        The function returns a list of people with associated Lodging, Per Sim, and Ground values.
        Example Output: [u'Dave Smith', 420.66, 36.6, 48.0]
        '''
        output_list = []
        local_lines = []
        names = []
        for line_item in self.line_items:
            if line_item.gl_in == 'Lodging' or line_item.gl_in == 'Per Diem' or line_item.gl_in == 'Ground':
                if line_item.person != 'Deposit':
                    local_lines.append(line_item)
                    self.remaining_line_items.remove(line_item)
                    if line_item.person not in names:
                        output_list.append([line_item.person,0.0,0.0,0.0])
                        names.append(line_item.person)

        for line_item in local_lines:
            if line_item.gl_in == 'Lodging':
                for record in output_list:
                    if line_item.person == record[0]:
                        record[1] = line_item.amount
            elif line_item.gl_in == 'Per Diem':
                for record in output_list:
                    if line_item.person == record[0]:
                        record[2] = line_item.amount
            elif line_item.gl_in == 'Ground':
                for record in output_list:
                    if line_item.person == record[0]:
                        record[3] = line_item.amount
        return output_list

    def coversheet_person_rows(self):
        output_list = []
        data = self.coversheet_per_person()
        for record in data:
            excel_sum_formula = '=sum(B%s:K%s)' % (self.coversheet_row, self.coversheet_row)
            row = [record[0], None, None,None,None,record[3],None,record[2],record[1],None,None, excel_sum_formula]
            output_list.append(row)
            self.coversheet_row += 1
        return output_list

    def coversheet_banquet(self):
        output_list = []
        local_lines = []
        names = []

        for line_item in self.remaining_line_items:
            print line_item.memo
            if line_item.comments != '' and line_item.person != 'Deposit':
                local_lines.append(line_item)
                self.remaining_line_items.remove(line_item)
                if line_item.comments not in names:
                    output_list.append([line_item.comments,0.0,0.0,0.0,0.0,0.0])
                    names.append(line_item.comments)


        for line_item in local_lines:

            for record in output_list:
                if line_item.comments == record[0]:
                    if line_item.gl_in == 'AudioVisual':
                        record[1] += line_item.amount
                    elif line_item.gl_in == 'Food & Beverage':
                        record[2] += line_item.amount
                    elif line_item.gl_in == 'Misc':
                        record[3] += line_item.amount
                    elif line_item.gl_in == 'Room Rental':
                        record[4] += line_item.amount
                    elif line_item.gl_in == 'Per Diem':
                        record[5] += line_item.amount


        return output_list



file = r"S:\CRG Internal Files\Financial Files\Hotel Bills 2017\NEK7324\NEK7324 breakdown.xlsx"
vendor = 'Ritz_Carlton'
invoice_date = '11.01.16'
program_manager = 'Regina LaRusso'
job_in = 7324
h = HotelFactory(file, vendor, invoice_date, job_in, program_manager)
#coversheet = CoversheetFactory(h)