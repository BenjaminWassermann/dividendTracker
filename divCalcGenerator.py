#!/usr/bin/env python3
# divCalcGenerator.py

# Creates a sheet with appropriate headers and formulas
# Fill with the list of tickers you would like to follow
# and save as 'dividendCalc.xlsx'

# import openpyxl
import openpyxl

# method encapsulation for CLI
def divCalcGenerator():

    # print working statement
    print('Please wait while I create a dividend calc template sheet...')

    # dicitonary containing cells and values
    entries = {'A1':'Stock ID',
               'B1':'Price',
               'C1':'Yield',
               'D1':'Annual Yield',
               'E1':'$price/$annual',
               'F1':'Annual yield for $1k',
               'G1':'Updated:'}

    # create new workbook and select active sheet
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = 'dividend calc'

    # place values in cells
    for position, entry in entries.items():
        sheet[position].value = entry
    

    # save sheet as toFill.xlsx and close
    book.save('toFill.xlsx')
    book.close()

    # print finished statement
    print('toFill.xlsx generated successfully')

# statement to run if not imported
if __name__=="__main__":
    divCalcGenerator()
