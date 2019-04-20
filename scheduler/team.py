#!/usr/bin/env python3
"""A module containing the Team class, for use in FLL tournament scheduling."""
from datetime import datetime, timedelta
class Team:
    """An FLL tournament team, storing a numeric ID, a name, a list of events, and a division."""
    def __init__(self, num, name, div=None):
        """Constructs a Team using the team's numeric id, name, and division (default=None)."""
        self.num = int(num)
        self.name = name
        self.div = div
        self.events = []

    def __str__(self):
        """Returns the str representation of the team. Does not include division."""
        return "Team " + str(self.num) + " " + str(self.name)

    def __repr__(self):
        """Returns a full representation of a team."""
        return "Team(num={}, name={}, div={}, events={}".format(self.num, self.name, self.div,
                                                                self.events)

    def info(self, with_div=False):
        """Returns a team's numeric ID, division (if true is passed to the function), and name."""
        return [self.num] + with_div*[self.div] + [self.name]

    def add_event(self, start_time, duration, activity_id, loc):
        """Adds an event at to the team's internal listing, then sorts the list of events by time.

        start_time -- the time the event starts (as a datetime)
        duration -- the duration of the event (as a timedelta)
        activity_id -- a numeric value representing the type of activity
        loc -- a numeric value representing the location the event will happen at"""
        self.events.append([start_time, duration, activity_id, loc])
        self.events.sort()

    def available(self, new_start, new_length, travel=timedelta(0)):
        """Returns true if the team is available for an new activity.

        new_start -- the time the new activity starts (as a datetime)
        new_length -- the duration of the new activity (as a timedelta)
        travel -- the travel time to allow between activies (as a timedelta, default 0)"""
        clear = [not (-e_length - travel < e_start - new_start < new_length + travel)
                 for (e_start, e_length, *others) in self.events]
        return all(clear)

    def next_event(self, time):
        """Returns the first event starting after time (if none, an event at datetime.max)."""
        return next((e for e in self.events if time < e[0]), (datetime.max, timedelta(0), -1, -1))

    def next_available(self, time, travel):
        return max((sum(e[:2], travel) for e in self.events
                   if (e[0] - travel) <= time <= sum(e[:2], travel)), default=time)

    def closest_events(self):
        """Returns the time between the two closest events scheduled for the team."""
        return min([self.events[i + 1][0] - self.events[i][0] - self.events[i][1]
                    for i in range(len(self.events) - 1)])
