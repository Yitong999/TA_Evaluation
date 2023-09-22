import numpy as np
from TA import TA
class TAs:
    def __init__(self):
        self.TAs_set = set() #contain TAs

    def add_survey(self, survey):
        for each in self.TAs_set:
            if each.name == survey.name:
                each.add_survey(survey)
                return

        # TA is not in the set yet
        new_ta = TA(survey.name)
        new_ta.add_survey(survey)
        self.TAs_set.add(new_ta)
        return

    def update_all(self):
        for each in self.TAs_set:
            each.update()




