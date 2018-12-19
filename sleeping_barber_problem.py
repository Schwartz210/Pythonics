'''
Djjkstra's Sleeping Barber Problem
Coded by Avi Schwartz on 10/26/18
'''
CHAIRS = 5  #Quantity of chairs in waiting room

class State(object):
    '''
    Main event-driven engine.
    '''
    vacant_chairs = CHAIRS
    now_serving = 1
    ticket = 1
    time = 0
    SLEEP = 0
    WORKING = 1
    FREE = 2
    status = SLEEP

    @staticmethod
    def execute():
        '''
        Controls main flow for a time increment
        '''
        while not AllCustomers.is_done():
            time_message = 'Time: ' + str(State.time)
            Status.to_console(0, time_message)
            customer_being_served = AllCustomers.get_customer_who_is_getting_their_haircut_currently()
            if customer_being_served:
                if customer_being_served.ticket == -1:
                    ticket_message = 'Customer being served never got a ticket because waiting room was empty'
            else:
                ticket_message = 'Ticket being served: ' + str(State.now_serving)
            Status.to_console(1, ticket_message)
            vacancy_message = 'Waiting room vacancy: ' + str(State.vacant_chairs)
            Status.to_console(1, vacancy_message)
            waiting_room_list_message = 'Customers in waiting room: ' + \
                                        AllCustomers.get_customers_with_status(Status.WAITING)
            Status.to_console(2, waiting_room_list_message)
            AllCustomers.arrival_handler()
            if State.status == State.WORKING:
                customer_being_served = AllCustomers.get_customer_who_is_getting_their_haircut_currently()
                customer_being_served.process_time -= 1
                customer_being_served.is_finished()
            elif State.status == State.FREE:
                AllCustomers.pick_next_for_haircut()
            elif State.status == State.SLEEP:
                if not AllCustomers.is_done():
                    AllCustomers.pick_next_for_haircut()
            State.time += 1
        Status.to_console(1, 'All customers have been serviced')
        Status.to_console(1, 'Summary')
        completed_summary_message = 'These customers successfully got their hair cut: ' + \
                                    AllCustomers.get_customers_with_status(Status.COMPLETE)
        Status.to_console(2, completed_summary_message)
        walkout_summary_message = 'These customers walked out: ' + AllCustomers.get_customers_with_status(Status.WALK_OUT)
        Status.to_console(2, walkout_summary_message)


class Status(object):
    '''
    Helper class. Holds constants which describe state of each customer.
    '''
    NOT_YET_ARRIVED = 0
    WAITING = 1
    HAIRCUT_IN_PROGRESS = 2
    COMPLETE = 3
    WALK_OUT = 4

    @staticmethod
    def to_console(indent_qty, message):
        '''
        To aid in printing messages to the console at specific levels of indentation
        '''
        indent = '    '
        text = indent * indent_qty
        text += message
        print(text)


class Customer(object):
    '''
    A single customer unit. Has name, arrival_time, process_time. Manages state for a single job
    '''
    def __init__(self, id, arrival_time, haircut_process_time):
        self.id = id
        self.arrival_time = arrival_time
        self.process_time = haircut_process_time
        self.ticket = -1
        self.status = Status.NOT_YET_ARRIVED

    def arrive_at_shop(self):
        '''
        Triggered whem customer arrives at shop. This method represents the how the customer decides if its time to
        take a ticket, get their hair cut, or walkout.
        '''
        def is_there_vacancy():
            '''
            Determines if there are vacant chairs in waiting room. Returns bool
            '''
            if State.vacant_chairs > 0:
                return True
            else:
                return False

        def take_a_number():
            '''
            Issues the customer a waiting ticket if they must wait.
            '''
            message = 'Customer %s is taking a ticket. Ticket number: %s' % (self.id, State.ticket)
            Status.to_console(2, message)
            self.ticket = State.ticket
            State.ticket += 1
            self.status = Status.WAITING

        has_arrived_message = 'Customer %s has arrived.' % self.id
        Status.to_console(1, has_arrived_message)
        if State.status == State.SLEEP:  #Checks if barber is sleeping
            State.status = State.WORKING
            Status.to_console(1, 'The Barber is now WORKING')
            self.status = Status.HAIRCUT_IN_PROGRESS
        else:
            if is_there_vacancy():
                vacancy_message = 'VACANCY: TRUE (availability=%s)' % State.vacant_chairs
                Status.to_console(1, vacancy_message)
                take_a_number()
                State.vacant_chairs -= 1
            else:  #Walking out
                self.status = Status.WALK_OUT
                Status.to_console(1, 'VACANCY: FALSE')
                walkout_message = 'Customer %s is walking out...' % self.id
                Status.to_console(2, walkout_message)

    def is_finished(self):
        '''
        This method handles process completion and continuation
        '''
        if self.process_time == 0:   #completion
            self.status = Status.COMPLETE
            transition_to_complete_message = 'Customer %s status=COMPLETE' % self.id
            Status.to_console(1, transition_to_complete_message)
            State.status = State.FREE
            Status.to_console(1, 'Barber is now FREE')
            if self.ticket != -1:
                State.now_serving += 1
            now_service_message = 'Now serving ticket: ' + str(State.now_serving)
            Status.to_console(1, now_service_message)
        else:  #continuation
            no_transition_message = 'Customer %s is still getting his haircut. Time remaining=%s' % \
                                    (self.id, self.process_time)
            Status.to_console(1, no_transition_message)


class AllCustomers(object):
    '''
    Collection data structure. Holds all customers. Contains methods that traversing entire customer list
    '''
    CUSTOMERS = [   #Sample data
        Customer('A', 1, 2),
        Customer('B', 2, 4),
        Customer('C', 4, 2),
        Customer('D', 5, 3),
        Customer('E', 9, 2),
        Customer('F', 12, 1),
        Customer('G', 15, 2),
        Customer('H', 16, 5),
        Customer('I', 16, 5),
        Customer('J', 17, 2),
        Customer('K', 20, 1),
        Customer('L', 21, 2),
        Customer('M', 22, 2),
        Customer('N', 23, 5),
        Customer('O', 24, 3),
    ]

    @staticmethod
    def is_done():
        '''
        Evaluates if all customers have been processed
        '''
        for customer in AllCustomers.CUSTOMERS:
            if customer.status in [0, 1, 2]:
                return False
        return True

    @staticmethod
    def arrival_handler():
        '''
        Determines which customers have arrived at the shop at that specific juncture in time
        '''
        for customer in AllCustomers.CUSTOMERS:
            if customer.arrival_time == State.time:
                customer.arrive_at_shop()

    @staticmethod
    def pick_next_for_haircut():
        '''
        Determines which customer will go next
        '''
        finding_next_customer_message = 'Finding the customer with ticket: ' + str(State.now_serving)
        Status.to_console(1, finding_next_customer_message)
        for customer in AllCustomers.CUSTOMERS:
            customer_ticket_message = 'Customer %s has ticket %s' % (customer.id, customer.ticket)
            Status.to_console(2, customer_ticket_message)
            if customer.ticket == State.now_serving:   #evaluates if customer's ticket is being served
                customer.status = Status.HAIRCUT_IN_PROGRESS
                message = 'Customer %s has sat in the barber chair' % customer.id
                Status.to_console(1, message)
                Status.to_console(1, 'A waiting room chair has become vacant')
                State.vacant_chairs += 1
                vacancy_message = 'Vacant chairs: ' + str(State.vacant_chairs)
                Status.to_console(2, vacancy_message)
                State.status = State.WORKING
                Status.to_console(1, 'Barber is WORKING')
                return

    @staticmethod
    def get_customer_who_is_getting_their_haircut_currently():
        '''
        Returns customer who is current getting their haircut
        '''
        for customer in AllCustomers.CUSTOMERS:
            if customer.status == Status.HAIRCUT_IN_PROGRESS:
                return customer

        return None

    @staticmethod
    def get_customers_with_status(status):
        '''
        returns a string list of all customers' IDs with a status that is specified by the argument
        '''
        out = ''
        counter = 0
        for customer in AllCustomers.CUSTOMERS:
            if customer.status == status:
                out += customer.id + ','
                counter += 1
        out = out[:-1]
        if counter == 0:
            out = 'No customers with this status.'
        return out


State.execute()