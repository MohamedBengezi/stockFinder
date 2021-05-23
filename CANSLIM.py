from pandas_datareader import data as pdr
from yahoo_fin import stock_info as si
from pandas import ExcelWriter
import yfinance as yf
import pandas as pd
import requests
import datetime
import time
import finnhub

finnhub_client = finnhub.Client(api_key='')

def convertStrToNum(num):
    if num[-1:] == 'K':  # Check if the last digit is K
        # Remove the last digit with [:-1], and convert to int and multiply by 1000
        return float(num[:-1]) * 1000
    elif num[-1:] == 'M':  # Check if the last digit is M
        # Remove the last digit with [:-1], and convert to int and multiply by 1000000
        return float(num[:-1]) * 1000000
    elif num[-1:] == 'B':  # Check if the last digit is M
        return float(num[:-1]) * 1000000000


yf.pdr_override()

stocklist = si.tickers_sp500()
index_name = '^GSPC'  # S&P 500

final = []
index = []
n = -1

exportList = pd.DataFrame(columns=['Stock', "RS_Rating", "50 Day MA",
                                   "150 Day Ma", "200 Day MA", "52 Week Low", "52 week High"])
ticker = input("Enter ticker: ")

#Current Quarterly Earnings
history = si.get_earnings_history(ticker)
eps_list = []
for i in range(min(len(history), 16)):
    dict1 = {}
    current_eps = history[i]['epsactual']
    date = history[i]['startdatetime']
    try:
        prev_eps = history[i+4]['epsactual']
        growth = (current_eps - prev_eps)/prev_eps * 100
        growth = int(growth)
    except:
        growth = 'N/A'
    if growth == 'N/A': continue
    dict1.update({'ticker': ticker, "EPS": current_eps, "% Change": growth, "Date": date[:10]})
    eps_list.append(dict1)
earnings_growth = pd.DataFrame(eps_list)

#Annual Earnings
earnings = si.get_earnings(ticker)['yearly_revenue_earnings']
earnings_list = []
index = 0

for i in range(min(25, len(history))):
    dict1 = {}
    date = history[index]['startdatetime'][:4]
    month = history[index]['startdatetime'][5:7]
    annual_eps = history[index]['epsactual']

    if month != '01' or annual_eps == None:
        index+= 1 
        continue

    for j in range(1, 4):
        annual_eps += history[index+j]['epsactual']
    index += 4
    
    dict1.update({'ticker': ticker, 'Year': int(date) - 1, 'Annual Earnings': annual_eps})
    earnings_list.append(dict1)

for i in range(len(earnings_list)-1):
    curr = earnings_list[i]['Annual Earnings']
    prev = earnings_list[i+1]['Annual Earnings']
    pct_chg = (curr - prev)/prev * 100
    earnings_list[i]['% Change'] = pct_chg
annual_earnings_growth = pd.DataFrame(earnings_list)



# New Products, Management, or New Highs
start = (datetime.datetime.now() - datetime.timedelta(weeks=2)).strftime('%Y-%m-%d')
end = datetime.datetime.now().strftime('%Y-%m-%d')
news = finnhub_client.company_news(ticker, _from=start, to=end)
news_sentiment = finnhub_client.news_sentiment(ticker)
pattern_recognition_days = finnhub_client.pattern_recognition(ticker, 'D')
pattern_recognition_weeks = finnhub_client.pattern_recognition(ticker, 'W')
support_resistance_days = finnhub_client.support_resistance(ticker, 'D')
support_resistance_weeks = finnhub_client.support_resistance(ticker, 'W')





print()
print(earnings_growth)
print(annual_earnings_growth)
print("News: ", news, '\n')
print("Sentiment: ", news_sentiment, '\n')
print("Pattern recognition (days): ", pattern_recognition_days, '\n')
print("Pattern recognition (weeks): ", pattern_recognition_weeks, '\n')
print("Support/Resistance Levels (days): ", support_resistance_days)
print("Support/Resistance Levels (weeks): ", support_resistance_weeks)

writer = ExcelWriter("ScreenOutput.xlsx")
exportList.to_excel(writer, "Sheet1")
writer.save()
