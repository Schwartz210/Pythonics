from tabulate import tabulate

def bold(string):
    return '\033[1m' + string + '\033[0m'

class ErrorChecker(object):
    '''
    This class is designed to find errors in the data to prevent them from going into QB
    '''
    def __init__(self, transaction_list, trans_type):
        self.transaction_list = transaction_list
        self.trans_type = trans_type

    def balance(self):
        '''
        This method compares the value of a transaction with the sum of its line items in order to determine that they
        are in balance.
        '''
        error_count = 0
        print('=================   ERROR CHECK: Balances   =================')
        for transaction in self.transaction_list:
            tran_amount = transaction.amount
            values_of_line_items = []
            for line_item in transaction.line_items:
                values_of_line_items.append(line_item.amount)
            sum_of_line_items = sum(values_of_line_items)
            bal = tran_amount - sum_of_line_items

            str_bal = str(bal)
            if str_bal != '0.0':
                print(bal, tran_amount, sum_of_line_items)
                error_count += 1
        print('Error count: %s' % (error_count))

    def error_checker_missing_jobs(self):
        '''
        This method evaluates line items to make sure that anything getting posted to a COGS account has a job
        '''
        error_count = 0
        print('=================   ERROR CHECK: Job data   =================')
        for transaction in self.transaction_list:
            for line_item in transaction.line_items:
                if not line_item.job_out and line_item.gl_out[:7] == 'Meeting':
                    print 'ERROR: Missing job info for %s' % (line_item.job_in)
                    error_count += 1
        print('Error count: %s' % (error_count))

    def error_checker_missing_gl(self):
        '''
        This method evaluates line items to make sure that all have G/L's.
        '''
        error_count = 0
        print('=================   ERROR CHECK: Missing G/L   =================')
        for transaction in self.transaction_list:
            for line_item in transaction.line_items:
                if line_item.gl_out == None:
                    print('ERROR: G/L issue on line %s' % (line_item.line_num))
                    error_count += 1
        print('Error count: %s' % (error_count))

    def missing_program_date(self):
        error_count = 0
        print('=================   ERROR CHECK: Missing Program Date   =================')
        for transaction in self.transaction_list:
            for line_item in transaction.line_items:
                if line_item.program_date == None or line_item.program_date == '':
                    print 'ERROR: Missing program date on line %s' % (line_item.line_num)
                    error_count += 1
        print('Error count: %s' % (error_count))

    def missing_last_name(self):
        '''
        This method looks for missing last names for HCP's on check runs.
        '''
        error_count = 0
        print('=================   ERROR CHECK: Missing Last Name   =================')
        for transaction in self.transaction_list:
            if transaction.vendor_obj.last_name == None or transaction.vendor_obj.last_name == '':
                print 'ERROR: Missing vendor last name on line %s' % (transaction.line_num)
                error_count += 1
        print('Error count: %s' % (error_count))

    def missing_link(self):
        '''
        This method looks for missing document links for checks and invoices.
        '''
        error_count = 0
        print('=================   ERROR CHECK: Missing Links   =================')
        for transaction in self.transaction_list:
            if transaction.link == None or transaction.link == '':
                print 'ERROR: Missing document link on line %s' % (transaction.line_num)
                error_count += 1
        print('Error count: %s' % (error_count))

    def missing_vendor(self):
        error_count = 0
        print('=================   ERROR CHECK: Missing Vendors   =================')
        for transaction in self.transaction_list:
            if transaction.vendor_obj == None:
                print 'ERROR: Missing vendor object on line %s' % (transaction.line_num)
                error_count += 1
        print('Error count: %s' % (error_count))
        if error_count > 0:
            print('Possible solution: Run new address report from QuickBooks')

    def print_transactions(self):
        '''
        This methods prints all transaction/line item data, and runs all error checks
        '''
        for transaction in self.transaction_list:
            grid_data = []
            splt_data = []
            grid_data.append(transaction.tabulate_box)
            print('========================================================================================================================')
            print tabulate(grid_data,headers=[bold('Vendor'),bold('Date'),bold('Invoice Num'),bold('Amount')])

            for line_item in transaction:
                splt_data.append(line_item.tabulate_box)
            print tabulate(splt_data,headers=[bold('G/L'),bold('Amount'),bold('Memo'),bold('Job')])
            print transaction.link
            print('\n\n\n\n\n\n')

        print('========================================================================================================================')
        self.error_checker_missing_jobs()
        self.balance()
        self.error_checker_missing_gl()
        self.missing_link()
        self.missing_vendor()
        if self.trans_type != 'allocation_spread' and self.trans_type != 'invoice_list':
            self.missing_program_date()
            self.missing_last_name()
        print('========================================================================================================================')
