from Survey import Survey
import os
import numpy as np
import openpyxl

def read_excel():
    # TODO: iterate through all files in **Work_Space** folder
    dataframe = openpyxl.load_workbook('Work_Space/book.xlsx')

    dataframe = dataframe.active

    # Iterate the loop to read the cell values
    for row in range(0, dataframe.max_row):
        for col in dataframe.iter_cols(1, dataframe.max_column):
            print(col[row].value, end=' ')
        print()

if __name__ == '__main__':
    read_excel()
    survey_1 = Survey('Tom', [1, 4, 5, 3, 4, 5], 'excellent')

    print(survey_1.scores)