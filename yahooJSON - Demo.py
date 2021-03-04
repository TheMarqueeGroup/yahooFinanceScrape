# -*- coding: utf-8 -*-
"""
Program creates a quick comps table using live data from Yahoo Finance

"""
#%% Import Packages
import urllib.request #package to connect to website
import json #package to manipulate the JSON data
import pandas as pd

#%% JSON for Yahoo Finance
#function that grabs a Yahoo Finance JSON URL and outputs the results as a dictionary
def fnYFinJSON(stock):
	urlData = "https://query2.finance.yahoo.com/v7/finance/quote?symbols="+stock
	webUrl = urllib.request.urlopen(urlData)
	if (webUrl.getcode() == 200):
		data = webUrl.read()
	else:
	    print ("Received an error from server, cannot retrieve results " + str(webUrl.getcode()))
	yFinJSON = json.loads(data)
	return yFinJSON["quoteResponse"]["result"][0]
	
#%% Create a comps table based on tickers and fields needed
tickers = ['AAPL', 'MSFT', 'TSLA', 'BA', 'FB', 'AMZN', 'NFLX']
fields = {'shortName':'Company', 'bookValue':'Book Value', 'currency':'Curr',
		  'fiftyTwoWeekLow':'52W L', 'fiftyTwoWeekHigh':'52W H',
		  'regularMarketPrice':'Price',
		  'regularMarketDayHigh':'High', 'regularMarketDayLow':'Low',
		  'priceToBook':'P/B', 'trailingPE':'LTM P/E'}

results = {}
for ticker in tickers:
	tickerData = fnYFinJSON(ticker)
	singleResult = {}
	for key in fields.keys():
		if key in tickerData:
			singleResult[fields[key]] = tickerData[key]
		else:
			singleResult[fields[key]] = "N/A"
	results[ticker] = singleResult

#%% Results as DataFrame
#default is keys are columns
dfTransp = pd.DataFrame.from_dict(results)

#unless you set orient as index
df = pd.DataFrame.from_dict(results, orient='index')
df.to_excel("OutputYFin.xlsx")
