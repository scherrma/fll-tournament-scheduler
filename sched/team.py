class Team:
    def __init__(self, num, name):
        self.team_num = str(num)
        self.team_name = name
        self.events = []

    def __str__(self):
        return "Team " + self.team_num + " " + self.team_name

    def __repr__(self):
        return str(self)
    
if __name__ == "__main__":
    from datetime import datetime
    t = Team(0, 'Acrobatic Whales')
    t.judging[0] = datetime(1, 1, 1, 9)
    t.judging[1] = datetime(1, 1, 1, 10, 30)
    t.judging[2] = datetime(1, 1, 1, 11, 30)
 #   t.matches.append(datetime(1, 1, 1, 9, 45))
 #   t.matches.append(datetime(1, 1, 1, 11))
 #   t.matches.append(datetime(1, 1, 1, 11, 55))
 #   t.matches.append(datetime(1, 1, 1, 12, 55))
 #   t.matches.append(datetime(1, 1, 1, 12, 35))
