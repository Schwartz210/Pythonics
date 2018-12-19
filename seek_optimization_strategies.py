'''
Simulation of magnetic disc seek & search optimization strategies.
The four strategies: FirstComeFirstServe, ShortestSeekTimeFirst, LOOK, Circular-LOOK
Coded by Avi Schwartz 10/29/18
'''
class Status(object):
    '''
    Used to communicate with console.
    '''
    @staticmethod
    def to_console(indent_qty, message):
        '''
        To aid in printing messages to the console at specific levels of indentation
        '''
        indent = '    '
        text = indent * indent_qty
        text += message
        print(text)

    @staticmethod
    def time_format(time):
        return str(time/10.0)


class Request(object):
    '''
    Represents a request to retrieve data from a specific location on the MagneticDisc
    '''
    NOT_YET_ARRIVED = 0
    WAITING = 1
    PROCESSING = 2
    COMPLETE = 3
    STATUS_DICTIONARY = {NOT_YET_ARRIVED: 'NOT_YET_ARRIVED',
                         WAITING: 'WAITING',
                         PROCESSING: 'PROCESSING',
                         COMPLETE: 'COMPLETE'}

    def __init__(self, id, arrival_time, track, sector):
        self.id = id
        self.arrival_time = arrival_time * 10
        self.track = track
        self.sector = sector
        self.status = Request.NOT_YET_ARRIVED
        self.completion_timestamp = 0


class AllRequests(object):
    '''
    A collection datastructure that holds all of the requests. All has methods used for iterating through all requests.
    '''
    REQUESTS = []
    complete = False

    @staticmethod
    def reset_requests():
        '''
        Resets the sample data every time a Strategy initializes
        '''
        Status.to_console(0, 'Reset Requests')
        AllRequests.REQUESTS = [
            Request('A', 0, 45, 0),
            Request('B', 23, 132, 6),
            Request('C', 25, 20, 2),
            Request('D', 29, 23, 1),
            Request('E', 35, 198, 7),
            Request('F', 45, 170, 5),
            Request('G', 57, 180, 3),
            Request('H', 83, 78, 4),
            Request('I', 88, 73, 5),
            Request('J', 95, 150, 7)
        ]

    @staticmethod
    def is_done():
        '''
        Determines if there is work to be done
        '''
        if len(AllRequests.get_requests_with_status(Request.PROCESSING)) == 1:
            return
        for request in AllRequests.REQUESTS:
            if request.status in [0, 1, 2]:
                return
        Status.to_console(2, 'All requests are COMPLETE')

        AllRequests.complete = True

    @staticmethod
    def arrival_handler(current_time):
        '''
        Checks for new arrivals and updates the request's status from NOT_YET_ARRIVED to WAITING
        '''
        Status.to_console(1, 'Checking arrivals...')
        for request in AllRequests.REQUESTS:
            if request.arrival_time == current_time:
                request.status = Request.WAITING
                message = 'The status of Request %s has changed from NOT_YET_ARRIVED to WAITING' % request.id
                Status.to_console(2, message)

    @staticmethod
    def get_requests_with_status(status):
        '''
        Returns a list of requests of a status specified by the argument
        '''
        out = []
        for request in AllRequests.REQUESTS:
            if request.status == status:
                out.append(request)
        return out


class MagneticDisc(object):
    '''
    This class simulates the magnetic disc.
    '''
    SECTOR = 1
    TRACK = SECTOR * 8
    SURFACE = TRACK * 200
    rotation_time = 70
    transfer_time = 12
    remaining_transer_time = transfer_time
    current_track = 0
    current_sector = 0
    current_request = None
    seek_time = 0.0
    search_time = 0.0
    working = False
    time = 0

    @staticmethod
    def calculate_seek_time(number_of_tracks):
        movement_constant = 100
        return movement_constant + 1 * abs(number_of_tracks)

    @staticmethod
    def calculate_search_time(number_of_sectors):
        return 8.75 * number_of_sectors

    @staticmethod
    def start_new_request():
        def set_search_and_seek_times():
            def set_seek_time():
                tracks = abs(MagneticDisc.current_request.track - MagneticDisc.current_track)
                traversal_message = 'The read head must traverse from track %s to track %s which is %s tracks' % \
                                    (MagneticDisc.current_track, MagneticDisc.current_request.track, tracks)
                Status.to_console(1, traversal_message)
                MagneticDisc.seek_time = MagneticDisc.calculate_seek_time(tracks)
                seektime_message = 'Seek time: ' + Status.time_format(MagneticDisc.seek_time)
                Status.to_console(2, seektime_message)

            def set_search_time():
                sectors = abs(MagneticDisc.current_request.sector - MagneticDisc.current_sector)
                traversal_message = 'Disc must rotate from sector %s to sector %s which is %s sectors' % \
                                    (MagneticDisc.current_sector, MagneticDisc.current_request.sector, sectors)
                Status.to_console(1, traversal_message)
                MagneticDisc.search_time = MagneticDisc.calculate_search_time(sectors)
                search_time_message = 'Search time: ' + Status.time_format(MagneticDisc.search_time)
                Status.to_console(2, search_time_message)

            set_seek_time()
            set_search_time()

        set_search_and_seek_times()
        MagneticDisc.remaining_transer_time = MagneticDisc.transfer_time

    @staticmethod
    def age():
        Status.to_console(1, 'Aging Report:')
        seek_message = 'Seek time remaining: ' + Status.time_format(MagneticDisc.seek_time)
        search_message = 'Search time remaining: ' + Status.time_format(MagneticDisc.search_time)
        Status.to_console(2, seek_message)
        Status.to_console(2, search_message)
        if MagneticDisc.seek_time < 0:
            MagneticDisc.seek_time = 0
        if MagneticDisc.search_time < 0:
            MagneticDisc.search_time = 0
        if MagneticDisc.seek_time <= 0 and MagneticDisc.search_time <= 0:
            if MagneticDisc.remaining_transer_time > 0:
                transfer_time_message = 'Transfer time remaining: ' + Status.time_format(MagneticDisc.remaining_transer_time)
                Status.to_console(2, transfer_time_message)
                MagneticDisc.remaining_transer_time -= 1
            else:
                MagneticDisc.current_track = MagneticDisc.current_request.track
                MagneticDisc.current_sector = MagneticDisc.current_request.sector
                MagneticDisc.current_request.status = Request.COMPLETE
                MagneticDisc.current_request.completion_timestamp = MagneticDisc.time
                MagneticDisc.working = False
                complete_message = 'Request %s is COMPLETE' % MagneticDisc.current_request.id
                Status.to_console(1, complete_message)
                current_track_message = 'Current Track: ' + str(MagneticDisc.current_track)
                current_sector_message = 'Current Sector: ' + str(MagneticDisc.current_sector)
                Status.to_console(2, current_track_message)
                Status.to_console(2, current_sector_message)
        elif MagneticDisc.seek_time == 0:
            MagneticDisc.search_time -= 1
        elif MagneticDisc.search_time == 0:
            MagneticDisc.seek_time -= 1
        else:
            Status.to_console(1, 'Track and sector traversal')
            MagneticDisc.seek_time -= 1
            MagneticDisc.search_time -= 1


class Strategy(object):
    def __init__(self):
        self.time = 0
        self.request_queue = []
        self.__str__()
        AllRequests.reset_requests()
        AllRequests.complete = False
        self.completion_times = []

    def cutoff(self):
        if self.time > 5000:
            exit()

    def __str__(self):
        message = 'Initiating ' + self.__class__.__name__
        Status.to_console(0, message)

    def new_request_set(self, request):
        MagneticDisc.current_request = request
        self.request_queue.remove(request)
        MagneticDisc.current_request.status = Request.PROCESSING
        MagneticDisc.working = True
        message = 'Disc is now serving Request ' + MagneticDisc.current_request.id
        Status.to_console(1, message)
        MagneticDisc.start_new_request()

    def set_current_request(self):
        '''
        Abstract
        '''
        pass

    def time_increment_tasks(self):
        self.queue_loader()
        MagneticDisc.time = self.time

    def queue_loader(self):
        waiting_requests = AllRequests.get_requests_with_status(Request.WAITING)
        for request in waiting_requests:
            if request not in self.request_queue:
                self.request_queue.append(request)

    def print_request_summary(self):
        Status.to_console(1, 'Request Summary:')
        for request in AllRequests.REQUESTS:
            if request.status == Request.COMPLETE:
                message = 'Request %s: %s(completion_timestamp=%s)' \
                          % (request.id, Request.STATUS_DICTIONARY[request.status], Status.time_format(request.completion_timestamp))
            else:
                message = 'Request %s: %s' % (request.id, Request.STATUS_DICTIONARY[request.status])
            Status.to_console(2, message)

    def store_results(self):
        for request in AllRequests.REQUESTS:
            self.completion_times.append((request.completion_timestamp-request.arrival_time)/10)

    def execute(self):
        while not AllRequests.complete:
            self.cutoff()
            time_message = 'Time:: ' + Status.time_format(self.time) + ' ms'
            Status.to_console(0, time_message)
            AllRequests.arrival_handler(self.time)
            self.time_increment_tasks()
            if MagneticDisc.working:
                MagneticDisc.age()
            else:
                Status.to_console(1, 'Disc is idle. Setting up next request...')
                self.set_current_request()
            self.time += 1
            self.print_request_summary()
            AllRequests.is_done()
        self.store_results()


class FirstComeFirstServe(Strategy):
    def __init__(self):
        Strategy.__init__(self)

    def set_current_request(self):
        if len(self.request_queue) > 0:
            self.new_request_set(self.request_queue[0])
        else:
            Status.to_console(1, 'No requests are waiting. Disc will remain idle until new request arrives...')


class ShortestSeekTimeFirst(Strategy):
    def __init__(self):
        Strategy.__init__(self)

    def get_track_distance(self, request):
        return abs(MagneticDisc.current_track - request.track)

    def print_distances_for_all_waiting(self):
        requests = AllRequests.get_requests_with_status(Request.WAITING)
        for request in requests:
            message = request.id + ':' + str(self.get_track_distance(request))
            Status.to_console(2, message)

    def set_current_request(self):
        def get_shortest_distance():
            shortest = self.request_queue[0]
            for request in self.request_queue:
                if self.get_track_distance(request) < self.get_track_distance(shortest):
                    shortest = request
            return shortest

        waiting_requests_message = 'WAITING requests: ' + str(len(AllRequests.get_requests_with_status(Request.WAITING)))
        Status.to_console(1, waiting_requests_message)
        if len(self.request_queue) > 0:
            Status.to_console(1, 'Summary of track distance:')
            self.print_distances_for_all_waiting()
            request = get_shortest_distance()
            self.new_request_set(request)
        else:
            Status.to_console(1, 'No requests are waiting. Disc will remain idle until new request arrives...')


class LOOK(Strategy):
    UP = 0
    DOWN = 1
    DIRECTION_DICT = {UP: 'UP', DOWN: 'DOWN'}

    def __init__(self):
        Strategy.__init__(self)
        self.direction = LOOK.UP

    def flip_direction(self):
        if self.direction == LOOK.UP:
            self.direction = LOOK.DOWN
        else:
            self.direction = LOOK.UP

    def get_order(self):
        def display(request_dictionary):
            summary_message = 'Summary of Distance(current track=%s, direction=%s):' % \
                              (MagneticDisc.current_track, LOOK.DIRECTION_DICT[self.direction])
            Status.to_console(1, summary_message)
            for key in request_dictionary.keys():
                request_message = 'Request %s: %s' % (key.id, request_dictionary[key])
                Status.to_console(2, request_message)

        def merge(same_direction, opposite_direction):
            temp = same_direction.copy()
            temp.update(opposite_direction)
            return temp

        def find_min(request_dictionary):
            minimum = request_dictionary.keys()[0]
            for key in request_dictionary.keys():
                if request_dictionary[key] < request_dictionary[minimum]:
                    minimum = key
            return minimum

        same_direction = {}
        opposite_direction = {}
        for request in self.request_queue:
            if self.direction == LOOK.UP:
                distance_to_edge = (MagneticDisc.SURFACE/MagneticDisc.TRACK) - MagneticDisc.current_track
                if request.track > MagneticDisc.current_track:
                    same_direction[request] = request.track - MagneticDisc.current_track
                else:
                    opposite_direction[request] = distance_to_edge + ((MagneticDisc.SURFACE/MagneticDisc.TRACK) - request.track)
            else:
                distance_to_edge = MagneticDisc.current_track
                if request.track < MagneticDisc.current_track:
                    same_direction[request] = request.track - MagneticDisc.current_track
                else:
                    opposite_direction[request] = distance_to_edge + request.track
        if len(same_direction) == 0:
            self.flip_direction()
        request_dictionary = merge(same_direction, opposite_direction)
        display(request_dictionary)
        minimum = find_min(request_dictionary)
        minimum_message = 'Request %s has the shortest distance of %s' % (minimum.id, request_dictionary[minimum])
        Status.to_console(3, minimum_message)
        return minimum


    def set_current_request(self):
        if len(self.request_queue) == 0:
            Status.to_console(1, 'No requests are waiting. Disc will remain idle until new request arrives...')
        elif len(self.request_queue) == 1:
            self.new_request_set(self.request_queue[0])
        else:
            request = self.get_order()
            self.new_request_set(request)


class CircularLOOK(Strategy):
    def __init__(self):
        Strategy.__init__(self)

    def set_current_request(self):
        def get_min(requests):
            minimum = requests[0]
            for request in requests:
                if request.track < minimum.track:
                    minimum = request
            return minimum

        def display(requests):
            summary_message = 'Track Summary(currently on Track %s):' % MagneticDisc.current_track
            Status.to_console(1, summary_message)
            for request in requests:
                request_message = 'Track %s: %s' % (request.id, request.track)
                Status.to_console(2, request_message)

        if len(self.request_queue) == 0:
            Status.to_console(1, 'No requests are waiting. Disc will remain idle until new request arrives...')
        elif len(self.request_queue) == 1:
            self.new_request_set(self.request_queue[0])
        else:
            requests = [request for request in self.request_queue if request.track > MagneticDisc.current_track]
            if len(requests) == 0:
                resetting_arm_message = 'No more requests to complete on this sweep. Resetting arm back from %s to Track 0' \
                                        % MagneticDisc.current_track
                Status.to_console(1, resetting_arm_message)
                MagneticDisc.current_track = 0
                requests = [request for request in self.request_queue if request.track > MagneticDisc.current_track]
            display(requests)
            request = get_min(requests)
            self.new_request_set(request)


def execute():
    def get_mean(results):
        total = 0
        for result in results:
            total += result
        return total / float(len(results))

    def get_variance(results):
        temp = 0
        mean = get_mean(results)
        for result in results:
            value = (result-mean) ** 2
            temp += value
        return temp / (len(results)-1)

    def string_format(class_name):
        new_string = str(class_name)
        for i in range(len(new_string), 21):
            new_string += ' '
        return new_string

    def display(results):
        def display_completion_times(results):
            Status.to_console(0, '\nCompletion Time Table:')
            for key in results.keys():
                completion_time_message = '%s: %s' % (key, results[key])
                Status.to_console(0, completion_time_message)

        def display_mean_table(results):
            Status.to_console(0, '\nMean Table:')
            for key in results.keys():
                mean_message = '%s:%s' % (key, get_mean(results[key]))
                Status.to_console(0, mean_message)

        def display_variance_table(results):
            Status.to_console(0, '\nVariance Table:')
            for key in results.keys():
                variance_message = '%s:%s' % (key, get_variance(results[key]))
                Status.to_console(0, variance_message)

        display_completion_times(results)
        display_mean_table(results)
        display_variance_table(results)

    strategies = [FirstComeFirstServe, ShortestSeekTimeFirst, LOOK, CircularLOOK]
    results = {}
    for strategy_class in strategies:
        strategy = strategy_class()
        strategy.execute()
        results[string_format(strategy_class.__name__)] = strategy.completion_times
    display(results)

execute()