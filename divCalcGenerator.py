#! Python3
# divCalcGenerator.py

# Creates a sheet with appropriate headers and formulas
# Fill with the list of tickers you would like to follow
# and save as 'dividend calc.xlsx'

# import openpyxl
import openpyxl

# print working statement
print('Please wait while I create a dividend calc template sheet...')

# dicitonary containing cells and values
entries = {'A1':'Stock ID',
           'B1':'Price',
           'C1':'Yield',
           'D1':'Recurrence',
           'E1':'Yield',
           'F1':'Annual Yield',
           'G1':'$stock/$div',
           'H1':'$stock/$annual',
           'I1':'Annual yield for $1k',
           'J1':'Updated:',
           'C2':'=F2/B2',
           'E2':'=F2/D2',
           'G2':'=B2/E2',
           'H2':'=B2/F2',
           'I2':'=FLOOR.MATH(1000/B2)*F2'}

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

