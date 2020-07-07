"""
Th
"""

class Evaluation:
    def __init__(self, filename):
        self.filename = filename
        self.evaluation_values = {
            'disclosed': 0,
            'quantitative_figure': {
                'non_monetary': 0,
                'monetary': 0
            }
        }
        self.marks_scheme = {
            'not_disclosed': 0,
            'one_less_disclosed': 1,
            'more_disclosed': 2,
            'one_quantitative': 3,
            'non_monetary_more_than_one': 4,
            'monetary': 5
        }




    def set_disclosed(self, sente):

    def get_disclosed(self):

    def set_non_monetary():

    def set_monetary_terms():

    def get_non_monetary_more_than_one(self):

    def get_monetary(self):
