from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import datetime
current_line = 750
MONTH = 'December 2017'


def next_line(number=None):
    global current_line
    if number:
        for _ in range(number):
            current_line -= 20
    else:
        current_line -= 20

def get_next_weekday(some_date, days_in_the_future):
    friday = 4
    if some_date.weekday() + days_in_the_future > friday:
        days_in_the_future = 7
    next_day = some_date + datetime.timedelta(days_in_the_future)
    return next_day


def pdf_coversheet_container():
    my_canvas = canvas.Canvas(r"H:\form.pdf", pagesize=letter)
    my_canvas.setLineWidth(.3)
    my_canvas.setFont('Helvetica', 12)
    my_canvas.drawString(30,current_line,'To:')
    next_line(2)
    my_canvas.drawString(30,current_line,'People')
    next_line()
    my_canvas.drawString(30,current_line,'More People')
    next_line(2)
    my_canvas.drawString(30,current_line,'Re:   ' + MONTH + ' Amex')
    next_line(2)
    my_canvas.drawString(30, current_line, 'Attached is your respective AMEX bill. I have also attached any receipts ')
    next_line()
    my_canvas.drawString(30, current_line, 'I have received in order to expedite your coding. ')
    next_line(2)
    my_canvas.drawString(75, current_line, 'Today:')
    today = datetime.datetime.now()
    my_canvas.drawString(225, current_line, today.strftime('%A, %B %d %Y'))
    next_line()
    return_by_date = get_next_weekday(today, 5)
    my_canvas.drawString(75, current_line, 'Please return by:')
    my_canvas.drawString(225, current_line, return_by_date.strftime('%A, %B %d %Y'))
    next_line()
    latest_date = get_next_weekday(return_by_date, 3)
    my_canvas.drawString(75, current_line, 'Or at Latest by:')
    my_canvas.drawString(225, current_line, latest_date.strftime('%A, %B %d %Y'))
    next_line(2)
    my_canvas.drawString(30,current_line,'Please, be specific as to job and what was purchased.')
    next_line(5)
    my_canvas.drawString(30,current_line,'As always, please be as detailed as possible in your descriptions of charges.')
    next_line(5)
    my_canvas.drawString(30,current_line,'Thank you for your cooperation!')
    next_line(2)
    my_canvas.drawString(30,current_line,'Finance')
    my_canvas.save()


pdf_coversheet_container()