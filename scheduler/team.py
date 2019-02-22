#!/usr/bin/env python3
import datetime

class Team:
    def __init__(self, num, name):
        self.num = int(num)
        self.name = name
        self.events = []

    def __str__(self):
        return "Team " + self.num + " " + self.name

    def __repr__(self):
        return str(self)

    def add_event(self, start_time, duration, name, loc):
        self.events.append([start_time, duration, name, loc])
        self.events.sort()

    def available(self, new_time, new_duration, travel=datetime.timedelta(0)):
        return all([new_time + new_duration + travel <= e_start or\
                e_start + e_duration + travel <= new_time for
                (e_start, e_duration, *others) in self.events])

    def next_event(self, time):
        try:
            return next((e for e in self.events if time < e[0]))
        except StopIteration:
            return (datetime.datetime.max, datetime.timedelta(0), -1, -1)
    
    def closest(self):
        return min([self.events[i + 1][0] - self.events[i][0] - self.events[i][1]
            for i in range(len(self.events) - 1)])
