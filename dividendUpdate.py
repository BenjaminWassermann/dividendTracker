#!/usr/bin/env python3
# dividendUpdate.py

# reads a list of tickers from an excel sheet
# updates current price and annual yield

# import statements for openpyxl, datetime, math, and yahoo_fin
import openpyxl, datetime, math
from yahoo_fin import stock_info as si

# encapsulate in method for command line calls
def dividendUpdate():

    # load dividend workbook and select active sheet
    divBook = openpyxl.load_workbook('dividendCalc.xlsx')
    divSheet = divBook.active

    # print start message
    print('Please wait while I update and re-calculate tickers...')

    # check and update each stock ticker
    for row in divSheet:

        # check to see if header or ticker
        if row[0].value != "Stock ID":

            # aquire the row id number
            rowID = row[0].row

            # generate strings for cols C, E, and F for rowID
            cRow = 'C%s' % rowID
            eRow = 'E%s' % rowID
            fRow = 'F%s' % rowID

            # generate strings for functions at C[rowID] and E[rowID]
            cFun = '=D%s/B%s' % (rowID, rowID)
            eFun = '=B%s/D%s' % (rowID, rowID)

            # place strings for functions at C[rowID] and E[rowID]
            divSheet[cRow].value = cFun
            divSheet[eRow].value = eFun    

            # store ticker in ticker
            ticker = row[0].value
            print(ticker)

            # query and store current price and data dictionary
            try:
                price = si.get_live_price(ticker)
                div = si.get_quote_table(ticker)['Forward Dividend & Yield'][:4]

                # update spreadsheet with new price and yield info
                row[1].value = price
                row[3].value = div

                # calculate shares per $1000
                shares = math.floor(1000/price)

                # generate function to calculate annual div yield per $1000 of shares
                fFun = '=%f * %f' % (float(shares), float(div))

                # places function
                divSheet[fRow].value = fFun

            except:
                print('There was a problem updating %s...' % ticker)
                continue

    # collect update date
    now = datetime.datetime.now()

    # place now string in K1
    divSheet['H1'].value = str(now)

    # save and close workbook
    divBook.save('dividendCalc.xlsx')
    divBook.close()

    # print end message
    print('Tickers updated!')

# statement to run if not imported
if __name__=="__main__":
    dividendUpdate()
            
