import numpy as np
class TA:
    def __init__(self, name):
        self.name = name
        self.scores_table = np.zeros((5, 5))
        self.comments = []
        self.total = np.zeros(5)
        self.avg = np.zeros(5)
        self.num_surveys = 0

    def add_survey(self, survey):
        if survey.name != self.name:
            return False

        scores = survey.scores
        for i in range(5):
            score = scores[i]
            self.scores_table[i][score - 1] += 1
            self.total[i] += 1
        self.comments.append(survey.comment)
        self.num_surveys += 1


    def update(self):
        for i in range(5):
            sum = 0
            for j in range(5):
                sum = sum + (j + 1) * self.scores_table[i][j]

            self.avg[i] = sum / self.total[i]


