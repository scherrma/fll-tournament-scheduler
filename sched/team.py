class Team:
    team_num = ''
    team_name = ''
    judging = (None, None, None)
    matches = None

    def __init__(self, num, name):
        self.team_num = str(num)
        self.team_name = name

    def __str__(self):
        return "Team " + self.team_num + " " + self.team_name

    def __repr__(self):
        return str(self)

