#! Python3
# dividendUpdate.py

# reads a list of tickers from an excel sheet
# updates current price and annual yield

# import statements for openpyxl and yahoo_fin
import openpyxl, datetime
from yahoo_fin import stock_info as si

# load dividend workbook and select active sheet
divBook = openpyxl.load_workbook('dividend calc.xlsx')
divSheet = divBook.active

# print start message
print('Please wait while I update and re-calculate tickers...')

# check and update each stock ticker
for row in divSheet:

    # check to see if header or ticker
    if row[0].value != "Stock ID":

        # store ticker in ticker
        ticker = row[0].value

        # query and store current price and data dictionary
        price = si.get_live_price(ticker)
        div = si.get_quote_table(ticker)['Forward Dividend & Yield'][:4]

        # update spreadsheet with new price and yield info
        row[1].value = price
        row[5].value = div

# collect update date
now = datetime.datetime.now()

# place now string in K1
divSheet['K1'].value = str(now)

# save and close workbook
divBook.save('dividend calc.xlsx')
divBook.close()

# print end message
print('Tickers updated!')

        
