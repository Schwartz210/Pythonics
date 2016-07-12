class Job(object):
    def __init__(self, packet):
        self.number = packet[0]
        self.full_name = packet[1]
        self.is_master = self.is_master_job()
        self.short_name = self.find_short_job()

    def is_master_job(self):
        '''
        Evualates if job is master job with multiple subjobs, returns boolean.
        Looks for 2 colons (":")
        '''
        if self.full_job.count(':') == 2:
            return True
        else:
            return False

    def standard_job(self):
        end = self.full_job.find(' ')
        result = self.full_job[0:end]
        return result

    def sub_job(self, full_job):
        '''
        Input: 'Job1234 Master Job:789 New York, NY 06/24/16:Ground'
        Output: 'Job1234-789'
        '''
        digits = 3
        colon = self.full_job.find(':') + 1
        digit_end = colon + digits
        program_num = self.full_job[colon:digit_end]
        master_job_end = self.full_job.find(' ')
        master_job = self.full_job[0:master_job_end]
        result = master_job + '-' + program_num
        return result

    def find_short_job(self):
        if self.is_master:
            return self.sub_job(self.full_job)
        else:
            return self.standard_job(self.full_job)

    def get_location(self):
        '''
        Input: 'Job1234 Master Job:789 New York, NY 06/24/16:Ground'
        Output: ('New York', 'NY')
        '''
        if self.is_master:
            length = len(self.full_job)
            colon = self.full_job.find(':')
            remove_left = colon + 5
            right_slice = self.full_job[remove_left:length]
            second_colon = right_slice.find(':')
            remove_right = second_colon - 9
            location_str = right_slice[0:remove_right]
            comma = location_str.find(',')
            city = location_str[0:comma]
            state = location_str[comma+2:]
            location = (city, state)
            return location
        else:
            return None