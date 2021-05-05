import numpy as np
import pandas as pd
import xlsxwriter
import requests
from scipy import stats
import math

stocks = pd.read_csv('sp_500_stocks.csv')
from secrets import IEX_CLOUD_API_TOKEN


#symbol = 'aapl'
#api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}'
#data = requests.get(api_url).json()

#price = data['latestPrice']
#pe_ratio = data['peRatio']

# batch call
def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]   
        
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

# building dataframe
my_columns = ['Ticker', 'Price', 'Price-to-Earnings Ratio', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns = my_columns)
for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    data[symbol]['quote']['peRatio'],
                    'N/A'
                ],
                index = my_columns
            ),
            ignore_index = True
        )

# sorting by lowest PE and removing negative PEs
final_dataframe.sort_values('Price-to-Earnings Ratio', ascending = True, inplace = True)
final_dataframe = final_dataframe[final_dataframe['Price-to-Earnings Ratio'] > 0]
final_dataframe = final_dataframe[:50]
final_dataframe.reset_index(inplace = True)
final_dataframe.drop('index', axis = 1, inplace = True)
#print(final_dataframe)

# calculating number of shares to buy
def portfolio_input():
    global portfolio_size
    portfolio_size = input("Enter the value of your portfolio:")

    try:
        val = float(portfolio_size)
    except ValueError:
        print("That's not a number! \n Try again:")
        portfolio_size = input("Enter the value of your portfolio:")

###portfolio_input()
###position_size = float(portfolio_size)/len(final_dataframe.index)
###for row in final_dataframe.index:
###    final_dataframe.loc[row, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[row, 'Price'])
#print(final_dataframe)

symbol = 'AAPL'
batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol}&types=quote,advanced-stats&token={IEX_CLOUD_API_TOKEN}'
data = requests.get(batch_api_call_url).json()
#print(data['AAPL']['advanced-stats'])
#Price-to-earnings ratio
pe_ratio = data[symbol]['quote']['peRatio']

#Price-to-book ratio
pb_ratio = data[symbol]['advanced-stats']['priceToBook']
#print(pb_ratio)

#Price-to-sales ratio
ps_ratio = data[symbol]['advanced-stats']['priceToSales']
#print(ps_ratio)

#Enterprise Value divided by Earnings Before Interest, Taxes, Depreciation, and Amortization (EV/EBITDA)
enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']
ebitda = data[symbol]['advanced-stats']['EBITDA']
ev_to_ebitda = enterprise_value / ebitda
#print(ev_to_ebitda)

#Enterprise Value divided by Gross Profit (EV/GP)
gross_profit = data[symbol]['advanced-stats']['grossProfit']
ev_to_gross_profit = enterprise_value / gross_profit
#print(ev_to_gross_profit)

# dataframe using robust value
rv_columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy',
    'Price-to-Earnings Ratio',
    'PE Percentile',
    'Price-to-Book Ratio',
    'PB Percentile',
    'Price-to-Sales Ratio',
    'PS Percentile',
    'EV/EBITDA',
    'EV/EBITDA Percentile',
    'EV/GP',
    'EV/GP Percentile',
    'RV Score'
]

rv_dataframe = pd.DataFrame(columns = rv_columns)
print(rv_dataframe)

for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote,advanced-stats&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']
        ebitda = data[symbol]['advanced-stats']['EBITDA']
        gross_profit = data[symbol]['advanced-stats']['grossProfit']

        try:
            ev_to_ebitda = enterprise_value / ebitda
        except TypeError:
            ev_to_ebitda = np.NaN

        try:
            ev_to_gross_profit = enterprise_value / gross_profit
        except TypeError:
            ev_to_gross_profit = np.NaN

        rv_dataframe = rv_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    'N/A',
                    data[symbol]['quote']['peRatio'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToBook'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToSales'],
                    'N/A',
                    ev_to_ebitda,
                    'N/A',
                    ev_to_gross_profit,
                    'N/A',
                    'N/A'
                ],
                index = rv_columns
            ),
            ignore_index = True
        )

#print(rv_dataframe)

# missing values
#print(rv_dataframe[rv_dataframe.isnull().any(axis=1)])
#print(rv_dataframe.columns)

for column in ['Price-to-Earnings Ratio', 'Price-to-Book Ratio', 'Price-to-Sales Ratio', 'EV/EBITDA', 'EV/GP']:
    rv_dataframe[column].fillna(rv_dataframe[column].mean(), inplace = True)
print(rv_dataframe[rv_dataframe.isnull().any(axis=1)])

# calculating value percentiles
from scipy.stats import percentileofscore as score
metrics = {'Price-to-Earnings Ratio': 'PE Percentile',
'Price-to-Book Ratio': 'PB Percentile',
'Price-to-Sales Ratio': 'PS Percentile',
'EV/EBITDA': 'EV/EBITDA Percentile',
'EV/GP': 'EV/GP Percentile'}

for metric in metrics.keys():
    for row in rv_dataframe.index:
        rv_dataframe.loc[row, metrics[metric]] = score(rv_dataframe[metric], rv_dataframe.loc[row, metric])
#print(rv_dataframe)

# calculating RV score
from statistics import mean

for row in rv_dataframe.index:
    value_percentiles = []
    for metric in metrics.keys():
        value_percentiles.append(rv_dataframe.loc[row, metrics[metric]])
    rv_dataframe.loc[row, 'RV Score'] = mean(value_percentiles)
#print(rv_dataframe)

# selecting 50 best value stocks
rv_dataframe.sort_values('RV Score', ascending = True, inplace = True)
rv_dataframe = rv_dataframe[:50]
rv_dataframe.reset_index(drop = True, inplace = True)
#print(rv_dataframe)

# calculating number of shares to buy
portfolio_input()
position_size = float(portfolio_size) / len(rv_dataframe.index)
for row in rv_dataframe.index:
    rv_dataframe.loc[row, 'Number of Shares to Buy'] = math.floor(position_size / rv_dataframe.loc[row, 'Price'])
#print(rv_dataframe)