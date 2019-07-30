README for dividendCalc python project
Created by Wasserdemon
Last updated 7/30/2019

The purpose of this project is to make it easier to track stock prices and dividend yields.

Using openpyxl and yahoo_fin, dividendUpdate reads in stock tickers (by default drawn from 'dividendCalc.xlsx'),
updating the price and dividend yield attributes (by default to 'dividendCalc.xlsx').

Using openpyxl, divCalcGenerator creates an excel workbook called 'toFill.xlsx' which is
formatted with appropriate headers and sheet names.

Using openpyxl, click, and yahoo_fin, dividendUpdateCLI provides an advanced interface for the
dividendCalc sheet

The sheet is formatted to calculate the best dividend values of included tickers. It achieves
this by comparing the purchase price of a given stock to the annual yield. We want to invest 
in those stocks which have the best dividend yield per dollar spent. The final field calculates
annual yield for $1k of a given product. This field calculates the number of shares that could
be purchased for $1k at the last price and multiplies that by the yield per share. A stock with
a good price/yield ratio may still be too expensive to yield sufficient growth through dividend
income. This attribute allows one to examine the effect of scale and could be changed to a larger
number for an individual with a larger account.

Folder Contents:

	divCalcGenerator.py  - stand-alone script which generates an excel sheet 
			       with appropriate headers. Simple CLI interface.

	dividendUpdate.py    - stand-alone script which updates the data for all tickers
			       in a sheet. Simple CLI interface.

	dividendUpdateCLI.py - not dependent on the other included scripts. Advanced 
			       CLI interface for div sheet.

	dividendCalc.xlsx    - default input and output sheet for most methods.

	test.xlsx	     - not referenced by any script, but provides an extant
			       excel workbook to use as an input or output.

	toFill.xlsx 	     - referenced by the divCalcGenerator script as the save
			       location.

Getting Started:

	Installing Libraries:

		using pip install, install the following libraries:
			openpyxl, yahoo_fin, and click

	Simple CLI guide:

		1. Navigate in your command prompt to the directory which contains the
		   scripts and workbooks.
		2. If you do not currently have a formatted sheet, run divCalcGenerator
		   from the command line.
		3. Manually fill toFill.xlsx with desired tickers then save and close.
		4. Run dividendUpdate from the command line.

	Advanced CLI guide:

		Usage: dividendUpdateCLI.py [OPTIONS] COMMAND [ARGS]...

		  Updates all tickers in file_in and saves to file_out

		Options:
		  -i, --file-in TEXT   [default: dividendCalc.xlsx]
		  -o, --file-out TEXT  [default: dividendCalc.xlsx]
		  --help               Show this message and exit.

		Commands:
  		add         Adds a new ticker to file_in and saves to file_out
  		addlist     Takes a list of tickers and adds them to file_in, saving to...
  		delete      Deletes an extant ticker from file_in and saves to file_out
  		deletelist  Takes a list of tickers and deletes examples from file_in and...
  		new         Creates a new formatted sheet with no tickers and saves to...

		!*The above is the result of sending dividendUpdateCLI -- help*!

		1. Options:

			a. -i or --file-in, defines file-in for any method. Default
			   is dividendCalc.xlsx. Do not use quotes.

			b. -o or --file-out, defines file-out for any method. Default
			   is dividendCalc.xlsx. Do not use quotes.

			c. --help summarizes all function in the CLI


		2. Default function updates all tickers in file_in and saves to file_out

		3. add adds a new ticker to file_in and saves to file_out
		
			a. argument is a single string, do not use quotes

		4. addlist adds a list of new tickers to file_in and saves to file_out

			a. argument is one or more stings, do not use quotes or commas

		5. delete removes and extant ticker from file_in and saves to file_out
			
			a. argument is a single string, do not use quotes

		6. deletelist removes a list of extant tickers from file_in and saves to file_out
		
			a. argument is one or more strings, do not use quotes or commas
			
