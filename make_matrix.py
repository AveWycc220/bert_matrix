import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter
import os

""" CONSTS """
PATH = os.path.dirname(os.path.abspath(__file__)) + '\\'


def find_excel_file():
    return openpyxl.load_workbook(PATH + '\\input\\' + os.listdir(path=f'{PATH}\\input')[0])


if __name__ == '__main__':
    input_wb = find_excel_file()
    sheet = input_wb[input_wb.sheetnames[0]]
    matrix = np.zeros((sheet.max_row-2, sheet.max_column-1))
    print(matrix.shape)
    for i in range(2, sheet.max_column+1):
        for j in range(3, sheet.max_row+1):
            matrix[j-3][i-2] = sheet.cell(row=j, column=i).value
    matrix = np.dot(matrix.transpose(), matrix)
    print(matrix)
    output_wb = openpyxl.Workbook()
    output_wb.title = "Матрица Берта"
    output_sheet = output_wb[output_wb.sheetnames[0]]
    for i in range(1, sheet.max_column+1):
        output_sheet.cell(row=i, column=1).value = sheet.cell(row=2, column=i).value
        column_widths = len(sheet.cell(row=2, column=i).value)
        output_sheet.column_dimensions[get_column_letter(i)].width = column_widths * 1.5
    for i in range(1, sheet.max_column+1):
        output_sheet.cell(row=1, column=i).value = sheet.cell(row=2, column=i).value
        column_widths = len(sheet.cell(row=2, column=i).value)
        output_sheet.column_dimensions[get_column_letter(i)].width = column_widths * 1.5
    for i in range(2, sheet.max_column + 1):
        for j in range(2, sheet.max_column + 1):
            output_sheet.cell(row=i, column=j).value = matrix[i-2][j-2]
    output_wb.save(filename=PATH + '\\output\\A_' + os.listdir(path=f'{PATH}\\input')[0])

