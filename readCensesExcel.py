#! python3
# -*- coding: utf-8 -*-
import os

__author__ = 'CaptShaw'

"""
    readCensusExcel.py - Tabulates population and number of censes tracts for
    each county.
"""
import openpyxl, pprint

path = r'C:\Users\Shaw\PycharmProjects\excelwork\example\censuspopdata.xlsx'
wb = openpyxl.load_workbook(path)
sheet = wb.active
countyData = {}

for row in range(2, sheet.max_row + 1):
    state = sheet['B' + str(row)].value
    county = sheet['C' + str(row)].value
    pop = sheet['D' + str(row)].value

    countyData.setdefault(state, {})
    countyData[state].setdefault(county, {'tracts': 0, 'pop': 0})

    countyData[state][county]['tracts'] += 1
    countyData[state][county]['pop'] += int(pop)

print('Writing results...')
with open(r'C:\Users\Shaw\PycharmProjects\excelwork\census2010.py', 'w') as resultFile:
    resultFile.write('allData = ' + pprint.pformat(countyData))
print('Done.')

