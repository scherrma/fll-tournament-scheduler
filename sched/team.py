from datetime import timedelta

class Team:
    def __init__(self, num, name):
        self.team_num = str(num)
        self.team_name = name
        self.events = []

    def __str__(self):
        return "Team " + self.team_num + " " + self.team_name

    def __repr__(self):
        return str(self)

    def add_event(self, start_time, duration, name, loc):
        self.events.append((start_time, duration, name, loc))
        self.events.sort()

    def available(self, new_time, new_duration=timedelta(0), travel=timedelta(0)):
        if new_duration < 0:
            return false
        return all([new_time + new_duration + travel < e_start or\
                e_start + e_duration + travel < new_time for
                (e_start, e_duration, *others) in self.events])
    
    def closest(self):
        return min([self.events[i + 1][0] - self.events[i][0] - self.events[i][1]
            for i in range(len(self.events) - 1)])
