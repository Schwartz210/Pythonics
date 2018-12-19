from time import sleep
HOLD = 0
READY = 1
RUNNING = 2
FINISHED = 3
job_list = []

class Job(object):
    def __init__(self, name, arrival, cpu_cycle):
        self.name = name
        self.arrival = arrival
        self.cpu_cycle = cpu_cycle
        self.status = HOLD #0=HOLD, 1=READY, 2=RUNNING, 3=FINISHED
        self.waiting_time = 0
        self.turnaround_time = 0
        self.finish_time = -1

    def age(self):
        print '    Aging Job %s...' % (self.name)
        self.cpu_cycle -= 1
        print '    Job %s has %s remaining cpu time' % (self.name, self.cpu_cycle)

    def running(self, time):
        self.status = RUNNING
        print '    Job %s is now RUNNING...' % (self.name)
        self.waiting_time = time - self.arrival

    def completion_handler(self, time):
        if self.cpu_cycle == 0:
            self.status = FINISHED
            print '    Job %s has been switched to FINISHED' % (self.name)
            self.finish_time = time
            self.turnaround_time = self.finish_time - self.arrival + 1


class SchedulingAlgorithm(object):
    def __init__(self):
        self.time = 0
        self.preemptive = False
        self.average_waiting_time = -1
        self.average_turnaround_time = -1
        print '%s scheduling algorithm initiating...' % (self.__class__.__name__)

    def arrival_handler(self):
        def get_arriving_job():
            for job in job_list:
                if job.arrival == self.time:
                    return job
            return None

        job = get_arriving_job()
        if job:
            job.status = READY
            print '    Job %s is now READY' % (job.name)

    def stop_running_job(self):
        running_job = self.get_jobs_with_status(2)
        if running_job:
            running_job[0].status = 1
            print '    Job %s is moved from RUNNING to READY' % (running_job[0].name)

    def not_done(self):
        for job in job_list:
            if job.status != FINISHED:
                return True
        return False

    def get_jobs_with_status(self, status):
        out = []
        for job in job_list:
            if job.status == status:
                out.append(job)
        return out

    def print_snapshot(self):
        def print_holds():
            holds = self.get_jobs_with_status(0)
            print '        HOLDS:'
            if holds:
                for job in holds:
                    print '            Job %s' % (job.name)

        def print_readies():
            readies = self.get_jobs_with_status(1)
            print '        READY:'
            for job in readies:
                print '            Job %s' % (job.name)

        def print_running():
            def data_validation(jobs):
                if len(jobs) > 1:
                    raise Exception(jobs)
            job = self.get_jobs_with_status(2)
            data_validation(job)
            print '        RUNNING:'
            if job:
                print '            Job %s' % (job[0].name)

        def print_finished():
            finished = self.get_jobs_with_status(3)
            print '        FINISHED:'
            for job in finished:
                print '            Job %s' % (job.name)

        print '    Status of all jobs'
        print_holds()
        print_readies()
        print_running()
        print_finished()

    def print_end(self):
        waiting_times = []
        turnaround_times = []
        print '    PROCESS SUMMARY for %s:' % (self.__class__.__name__)
        for job in job_list:
            print '        Job %s: Waiting time=%s, Turnaround time=%s' % (job.name, job.waiting_time, job.turnaround_time)
            waiting_times.append(job.waiting_time)
            turnaround_times.append(job.turnaround_time)
        self.average_waiting_time = sum(waiting_times) / len(waiting_times)
        self.average_turnaround_time = sum(turnaround_times) / len(turnaround_times)
        print '        Average Waiting Time: %s' % (self.average_waiting_time)
        print '        Average Turnaround Time: %s' % (self.average_turnaround_time)

    def get_job_to_process(self):
        def get_currently_running_job():
            def data_validation(jobs):
                if len(out) > 1:
                    for job in jobs:
                        print job
                    raise Exception('More than one job at a time.')
            out = []
            for job in job_list:
                if job.status == RUNNING:
                    out.append(job)
            data_validation(out)
            if len(out) > 0:
                return out[0]
            else:
                return None

        if get_currently_running_job():
            if not self.preemptive:
                return get_currently_running_job()
            else: #preemptive
                job_to_process = self.get_next_in_queue()
                if job_to_process.name != get_currently_running_job().name:
                    self.stop_running_job()
                    job_to_process.running(self.time)
                return job_to_process
        else:
            job_to_process = self.get_next_in_queue()
            job_to_process.running(self.time)
            return job_to_process

    def get_next_in_queue(self):
        '''
        Abstract method gets implemented via childred
        '''
        pass

    def main_sequence(self):
        while self.not_done():
            sleep(0)
            print 'TIME: ', self.time
            self.arrival_handler()
            job_in_progress = self.get_job_to_process()
            job_in_progress.age()
            self.print_snapshot()
            job_in_progress.completion_handler(self.time)
            self.time += 1
        print '    All jobs are FINISHED.'
        self.print_snapshot()
        self.print_end()


class FirstComeFirstServe(SchedulingAlgorithm):
    def __init__(self):
        SchedulingAlgorithm.__init__(self)
        self.preemptive = False
        self.main_sequence()

    def get_next_in_queue(self):
        queued_jobs = self.get_jobs_with_status(1)
        earliest_arriving = Job('DEFAULT', 9999999, 0)
        for job in queued_jobs:
            if job.arrival < earliest_arriving.arrival:
                earliest_arriving = job
        return earliest_arriving


class ShortestJobNext(SchedulingAlgorithm):
    def __init__(self):
        SchedulingAlgorithm.__init__(self)
        self.preemptive = False
        self.main_sequence()

    def get_next_in_queue(self):
        queued_jobs = self.get_jobs_with_status(1)
        shortest_job = Job('DEFAULT', 0, 99999999)
        for job in queued_jobs:
            if job.cpu_cycle < shortest_job.cpu_cycle:
                shortest_job = job
        return shortest_job


class ShortestRemainingTime(SchedulingAlgorithm):
    def __init__(self):
        SchedulingAlgorithm.__init__(self)
        self.preemptive = True
        self.main_sequence()

    def get_next_in_queue(self):
        ready_and_running_jobs = self.get_jobs_with_status(1) + self.get_jobs_with_status(2)
        shortest_remaining_time_job = Job('DEFAULT', 0, 99999999)
        self.print_remaining_times()
        for job in ready_and_running_jobs:
            if job.cpu_cycle < shortest_remaining_time_job.cpu_cycle:
                shortest_remaining_time_job = job
        return shortest_remaining_time_job

    def print_remaining_times(self):
        print '    Remaining times for READY jobs:'
        ready_and_running_jobs = self.get_jobs_with_status(1) + self.get_jobs_with_status(2)
        for job in ready_and_running_jobs:
            print '        Job %s: %s' % (job.name, job.cpu_cycle)


class RoundRobin(SchedulingAlgorithm):
    def __init__(self):
        SchedulingAlgorithm.__init__(self)
        self.preemptive = True
        self.time_quantum = 4
        self.cycles = self.time_quantum
        self.iterator = 1
        self.main_sequence()

    def get_queued(self):
        out = []
        for job in job_list:
            if job.status == 1 or job.status == 2:
                out.append(job)
        return out

    def get_next_in_queue(self):
        def circular_fetch(ready_and_running_jobs):
            array_size = len(ready_and_running_jobs)
            if self.iterator >= array_size:
                self.iterator = 0
            job_to_process = ready_and_running_jobs[self.iterator]
            self.iterator += 1
            return job_to_process
        if self.cycles == self.time_quantum:
            print '    Time Quantum interval. A new job will enter RUNNING state'
            self.cycles = 0
            ready_and_running_jobs = self.get_queued()
            job_to_run = circular_fetch(ready_and_running_jobs)
        else:
            print '    Not a time quantum interval. No job change until previous RUNNING job moved to FINISHED'
            if self.time == 0:
                job_to_run = self.get_jobs_with_status(1)[0]
            elif len(self.get_jobs_with_status(2)) == 1:
                job_to_run = self.get_jobs_with_status(2)[0]
            elif len(self.get_jobs_with_status(2)) == 0:
                ready_and_running_jobs = self.get_queued()
                job_to_run = circular_fetch(ready_and_running_jobs)
        self.cycles  += 1
        return job_to_run


def execute():
    def reset_job_list():
        global job_list
        job_list = [
            Job('A', 0, 16),
            Job('B', 3, 2),
            Job('C', 5, 11),
            Job('D', 9, 6),
            Job('E', 10, 1),
            Job('F', 12, 9),
            Job('G', 14, 4)
        ]

    def print_comparision_table(algorithms):
        print 'COMPARISION TABLE:'
        for algorithm in algorithms:
            print '    %s: Average Waiting Time=%s, Average Turnaround Time=%s' % (algorithm.__class__.__name__,
                                                                                   algorithm.average_turnaround_time,
                                                                                   algorithm.average_waiting_time)
    algorithms = []
    scheduling_algorithms_classes = [FirstComeFirstServe, ShortestJobNext, ShortestRemainingTime, RoundRobin]
    for scheduling_algorithm_class in scheduling_algorithms_classes:
        reset_job_list()
        algorithm = scheduling_algorithm_class()
        algorithms.append(algorithm)
    print_comparision_table(algorithms)


execute()
