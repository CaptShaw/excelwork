#! python3
# -*- coding: utf-8 -*-

__author__ = 'CaptShaw'

"""
    12.13.2 blankRowInserter
    创建一个程序 blankRowInserter.py，它接受两个整数和一个文件名字符串作为
    命令行参数。我们将第一个整数称为 N，第二个整数称为 M。
    程序应该从第 N 行开 始，在电子表格中插入 M 个空行。
"""
import  openpyxl

def blankRowInserter(n,m,file):
    '''
    :param m: insert m rows
    :param n: from row no. n
    :param file: target excel file
    :return:  None
    '''
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
    sheet.insert_rows(n,amount=m)
    # for rowOfCellObjects in sheet.rows:
    #     print(rowOfCellObjects)

    wb.save('eg_changed.xlsx')


if __name__ == '__main__':
    blankRowInserter(2,5,r'C:\Users\Shaw\PycharmProjects\excelwork\example\example.xlsx')