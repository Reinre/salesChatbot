
class SalesReport:
    """Class to pack the query data, to be shown on the message"""
    def __init__(self, goal_status, goal_color, dates, last_year, saless, goal, sanji, rokuji, kuji, changes, monthly_goal, currents, lefts, progress):
        self.goal_status = goal_status
        self.goal_color = goal_color
        self.dates = dates
        self.last_year = last_year
        self.saless = saless
        self.goal = goal
        self.sanji = sanji
        self.rokuji = rokuji
        self.kuji = kuji
        self.changes = changes
        self.monthly_goal = monthly_goal
        self.currents = currents
        self.lefts = lefts
        self.progress = progress
   