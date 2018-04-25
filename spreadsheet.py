__author__ = 'aschwartz - Schwartz210@gmail.com'
import xlrd
from xlwt import Workbook

def format_date(date):
    if len(str(date)) == 0:
        return None
    elif len(str(date)) == 8 or len(str(date)) == 10:
        return date
    elif date == 'TBD':
        return date
    else:
        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(date, 0)
        day = str(day)
        month = str(month)
        if len(day) == 1: day = '0' + day
        if len(month) == 1: month = '0' + month
        formatted_date = month + '/' + day + '/' + str(year)[2:]
        return formatted_date

def open_file(path, tab, rows):
    """
    Open and read an Excel file
    """
    content = []
    book = xlrd.open_workbook(path)
    if type(tab) == int:
        tab = book.sheet_by_index(tab)
    else:
        tab = book.sheet_by_name(tab)
    for i in range(rows):
        try:
            content.append(tab.row_values(i))  # read a row
        except:
            break
    return content

def screen_out_empty_rows(list):
    output_list = []
    for row in list:
        try:
            string = str(row[0])
            if len(string) > 0:
                output_list.append(row)
        except:
            continue
    return output_list

def screen_out_col(list, first_col, last_col):
    output_list = []
    for row in list:
        output_list.append(row[first_col:last_col+1])
    return output_list

def keep_these_columns(data, *columns):
    columns = list(columns)
    output_list = []
    for record in data:
        smaller_record = []
        for column in columns:
            smaller_record.append(record[column])
        output_list.append(smaller_record)
    return output_list

def keep_columns(data, columns):
    output_list = []
    for record in data:
        smaller_record = []
        for column in columns:
            smaller_record.append(record[column])
        output_list.append(smaller_record)
    return output_list

def cr_screen_out_content(list):
    result = screen_out_empty_rows(list)
    result = screen_out_col(result, 0, 9)
    return result

def check_request_pull():
    checkrequest_list = open_file('Mass check request gui.xlsx',0,100)
    checkrequest_list = cr_screen_out_content(checkrequest_list)
    return checkrequest_list

def invoice_pull():
    invoice_list = open_file('Mass invoice gui.xlsx',0,100)
    invoice_list = cr_screen_out_content(invoice_list)
    print('OP: Data pulled from Invoice worksheet')
    return invoice_list

def super_string(rows):
    new_rows = []
    for row in rows:
        new_row = []
        for cell in row:
            try:
                new_cell = str(cell.encode('utf-8'))
                new_row.append(new_cell)
            except:
                new_cell = str(int(cell))
                new_row.append(new_cell)
        new_rows.append(new_row)
    return new_rows

def format_one_job_number(row):
    try:
        string = str(int(row))
        if len(string) < 4:
            zeros_to_add = 4 - len(string)
            for _ in range(zeros_to_add):
                string += '0'
        row = string
        return row
    except:
        return row

def format_job_numbers(rows):
    for row in rows:
        row[0] = format_one_job_number(row[0])

def col_to_num(col_str):
    '''
    https://stackoverflow.com/questions/7261936/convert-an-excel-or-spreadsheet-column-letter-to-its-number-in-pythonic-fashion
    '''
    expn = 0
    col_num = -1
    col_str = col_str.upper()
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num

def is_letter(character):
    if type(character) == int:
        return False
    elif character.isalpha():
        return True
    else:
        return False

def convert_to_column_numbers(columns):
    out = []
    for col in columns:
        if is_letter(col):
            out.append(col_to_num(col))
        else:
            out.append(int(col))
    return out

def make_spreadsheet(rows, **kwargs):
    '''
    Converts list_of_lists to spreadsheet.xls
    '''
    book = Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")
    if 'headers' in kwargs:
        rows.insert(0, kwargs['headers'])
    for row in range(len(rows)):
        for cell in range(len(rows[row])):
            sheet1.write(row, cell, rows[row][cell])
    if 'title' in kwargs:
        book.save(kwargs['title'])
    else:
        book.save('H:\generic report.xls')


class PowerRead(object):
    def __init__(self, path, tab, **kwargs):
        self.kwargs = kwargs
        self.data = self.open_file(path, tab)
        self.exclude()
        self.only_include()
        self.columns()
        self.column_range()
        self.no_empty_rows()
        self.inventory = self.factory()

    def open_file(self, path, tab):
        if 'row_quantity' in self.kwargs:
            return open_file(path, tab, self.kwargs['row_quantity'])
        else:
            return open_file(path, tab, 1000)

    def exclude(self):
        if 'exclude' in self.kwargs:
            temp = []
            column = convert_to_column_numbers([self.kwargs['exclude'].keys()[0]])[0]
            value = [self.kwargs['exclude'].values()[0]][0]
            for record in self.data:
                if record[column] != value:
                    temp.append(record)
            self.data = temp

    def only_include(self):
        if 'only_include' in self.kwargs:
            temp = []
            column = convert_to_column_numbers([self.kwargs['only_include'].keys()[0]])[0]
            value = [self.kwargs['only_include'].values()[0]][0]
            for record in self.data:
                if record[column] == value:
                    temp.append(record)
            self.data = temp

    def columns(self):
        if 'columns' in self.kwargs:
            self.data = keep_columns(self.data, convert_to_column_numbers(self.kwargs['columns']))

    def column_range(self):
        if 'column_range' in self.kwargs:
            if 'columns' in self.kwargs:
                raise Exception('columns and column_range keywords cannot be used in tandem')
            else:
                column_range = convert_to_column_numbers(self.kwargs['column_range'])
                self.data = screen_out_col(self.data, column_range[0], column_range[1])

    def no_empty_rows(self):
        if 'no_empty_rows' in self.kwargs:
            if type(self.kwargs['no_empty_rows']) != bool:
                raise Exception('no_empty_rows can only be set to a boolean')
            elif self.kwargs['no_empty_rows']:
                self.data = screen_out_empty_rows(self.data)
            elif not self.kwargs['no_empty_rows']:
                pass

    def factory(self):
        out = []
        if 'factory' in self.kwargs:
            for record in self.data:
                my_class = self.kwargs['factory']
                instance = my_class(record)
                out.append(instance)
        return out

    def inventory_to_array(self):
        out = []
        for item in self.inventory:
            out.append(item.to_array())
        return out

    def __str__(self):
        for record in self.data:
            print(record)

    def print_keywords(self):
        keywords = ['row_quantity', 'columns', 'column_range', 'no_empty_rows','factory','exclude', 'only_include']
        print(keywords)





