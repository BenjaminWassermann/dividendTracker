README for dividendCalc python project

The purpose of this project is to make it easier to track stock prices and dividend yields.

Using openpyxl and yahoo_fin, dividendCalc.py reads in stock tickers (by default drawn from 'dividendCalc.xlsx'),
updating the price and dividend yield attributes (by default to 'dividendCalc.xlsx').

Using openpyxl, divCalcGenerator.py creates an excel workbook called 'toFill.xlsx' which is
formatted with appropriate headers and formulas in the first row.

The sheet is formatted to calculate the best dividend values of included tickers. It achieves
this by comparing the purchase price of a given stock to the annual yield. We want to invest 
in those stocks which have the best dividend yield per dollar spent. The final field calculates
annual yield for $1k of a given product. This field calculates the number of shares that could
be purchased for $1k at the last price and multiplies that by the yield per share. A stock with
a good price/yield ratio may still be too expensive to yield sufficient growth through dividend
income. This attribute allows one to examine the effect of scale and could be changed to a larger
number for an individual with a larger account.

