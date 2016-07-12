from datetime import datetime

class Logger(object):
    '''
    This class serves to gather data pertaining to this program's operations.
    '''
    def __init__(self, function_name, message):
        self.function_name = function_name
        self.message = message
        self.log_files_dict = {'emails' : 'logs/email log.txt',
                               'test' : 'logs/logger.txt',
                               'open_file' : 'logs/logger.txt'}
        self.selected_log_file = self.log_files_dict[self.function_name]
        self.now = datetime.now()
        self.log()
        self.__str__()

    def log(self):
        '''
        This method opens log file, appends message, closes it.
        '''
        file = open(self.selected_log_file, 'a')
        file.write(str(self.now))
        file.write(' - ' + self.message + '\n')
        file.close()

    def __str__(self):
        print '================================================================================'
        print 'Logger has been initiated.'
        print 'Message: ', self.message
        print 'Log file: ', self.selected_log_file
        print 'Time logged: ', self.now
        print '================================================================================'

