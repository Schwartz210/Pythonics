'''
This module reads an attendee report and exports line items
'''
from super_transaction import SuperTransaction
from spreadsheet import open_file, screen_out_empty_rows, format_date

class ReportInfo(object):
    def __init__(self):
        self.version = None
        self.total = None
        self.vendor = None
        self.job = None
        self.divisor = None
        self.date = None
        self.deposit = None
        self.room_rental = None
        self.audiovisual = None
        self.misc = None
        self.food_and_bev = None
        self.attendees = []

    def __str__(self):
        print 'Total: ', self.total
        print 'Vendor: ', self.vendor
        print 'Job: ', self.job
        print 'Divisor: ', self.divisor
        print 'Date: ', self.date
        print 'Deposit: ', self.deposit
        print 'Room Rental: ', self.room_rental
        print 'AudioVisual: ', self.audiovisual
        print 'Misc: ', self.misc
        print 'Food & Beverage: ', self.food_and_bev
        print 'Attendees: '
        for attendee in self.attendees:
            print '    ', attendee[0]

class AttendeeLineItem(SuperTransaction):
    def __init__(self, **kwargs):
        self.set_attributes(kwargs)

    def set_attributes(self, kwargs):
        if 'gl_in' in kwargs.keys():
            self.gl_in = kwargs['gl_in']
        if 'short_job' in kwargs.keys():
            self.short_job = kwargs['short_job']
        if 'amount' in kwargs.keys():
            self.amount = kwargs['amount']
        if 'memo' in kwargs.keys():
            self.memo = kwargs['memo']
        self.tran_class = ''
        self.full_job = self.find_full_job_name_from_short(self.short_job)
        try:
            self.job_out = self.full_job + ':' + self.gl_in
        except:
            if not self.full_job:
                raise Exception('Issue with job: ', self.short_job)
            elif not self.gl_in:
                raise Exception('issue with GL: ', self.gl_in)
        self.client = self.find_client(self.full_job)
        self.gl_out = self.create_gl_out(self.gl_in, self.client)

    def get_output(self):
        return [self.gl_out, self.amount, self.memo, self.job_out, self.tran_class]



def make_attendee_memo(info, record):
    indices = [0,0,0]
    if info.version == 'Attendee Report - Version 3.0':
        indices = [1,2,3]
    elif info.version == 'Version 4.0':
        indices = [0,1,2]
    if record[3]:
        person_type = str(record[3]) + str('-')
    else:
        person_type = ''
    memo = '%s-%s %s-F&B-%s-%s[%s]{%s}-%s' % (info.vendor, record[indices[0]], record[indices[1]], record[indices[2]], person_type, info.divisor, info.food_and_bev, info.date)
    return memo

def read_report(filepath):
    data = open_file(filepath, 0, 1000)
    return screen_out_empty_rows(data)

def extract_info(data):
    info = ReportInfo()
    info.version = data[0][5]
    if info.version == 'Attendee Report - Version 3.0':
        info.total = data[2][1]
        info.vendor = data[3][1]
        info.job = data[4][1]
        info.divisor = data[5][1]
        info.date = format_date(data[6][1])
        info.deposit = data[8][1]
        info.room_rental = data[9][1]
        info.audiovisual = data[10][1]
        info.misc = data[11][1]
        info.food_and_bev = data[12][1]
        for record in data[14:]:
            memo = make_attendee_memo(info, record)
            items = record[1:5]
            items.append(memo)
            info.attendees.append(items)
    elif info.version == 'Version 4.0':
        info.total = data[2][1]
        info.vendor = data[3][1]
        info.job = data[4][1]
        info.date = format_date(data[5][1])
        info.deposit = data[7][1] * -1
        info.room_rental = data[8][1]
        info.audiovisual = data[9][1]
        info.misc = data[10][1]
        info.food_and_bev = data[11][1]
        info.divisor = len(data[13:])
        for record in data[13:]:
            memo = make_attendee_memo(info, record)
            items = record[0:4]
            items.append(memo)
            info.attendees.append(items)
    return info

def allocate(splits, total_amount):
    divisor = len(splits)
    allocation = round(total_amount // divisor)
    allocated_amount = allocation * divisor
    remainder = total_amount - allocated_amount
    for thing in splits:
        thing[1] += allocation
    iterator = 0
    while remainder > 0.01:
        if iterator == divisor:
            iterator = 0
        splits[iterator][1] += 0.01
        remainder -= 0.01
        iterator += 1
    for split in splits:
        split[1] = round(split[1], 2)
    return splits

def build_top_line_items(info):
    line_items = []
    if info.deposit:
        deposit = AttendeeLineItem(gl_in='Food & Beverage', short_job=info.job, amount=info.deposit)
        deposit.memo = '%s - Less Deposit - %s' % (info.vendor, info.date)
        line_items.append(deposit.get_output())
    if info.room_rental:
        room = AttendeeLineItem(gl_in='Room Rental', short_job=info.job, amount=info.room_rental)
        room.memo = '%s - Room Rental - %s' % (info.vendor, info.date)
        line_items.append(room.get_output())
    if info.audiovisual:
        av = AttendeeLineItem(gl_in='AudioVisual', short_job=info.job, amount=info.audiovisual)
        av.memo = '%s - AudioVisual - %s' % (info.vendor, info.date)
        line_items.append(av.get_output())
    if info.misc:
        misc = AttendeeLineItem(gl_in='Misc', short_job=info.job, amount=info.misc)
        misc.memo = '%s - Misc - %s' % (info.vendor, info.date)
        line_items.append(misc.get_output())
    return line_items

def build_attendee_line_items(info):
    line_items = []
    for attendee in info.attendees:
        a = AttendeeLineItem(gl_in='Food & Beverage', short_job=info.job, amount=0, memo=attendee[4])
        a.tran_class = attendee[1] + ', ' + attendee[0]
        line_items.append(a.get_output())
    return allocate(line_items, info.food_and_bev)

def build_line_items(info):
    line_items = build_top_line_items(info)
    attendee_lines = build_attendee_line_items(info)
    for line in attendee_lines:
        line_items.append(line)
    return line_items

def get_line_items_from_attendee_reports(filepath):
    data = read_report(filepath)
    info = extract_info(data)
    line_items = build_line_items(info)
    return line_items


