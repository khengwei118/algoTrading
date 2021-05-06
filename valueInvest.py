import numpy as np
import pandas as pd
import xlsxwriter
import requests
from scipy import stats
import math
from statistics import mean
from secrets import IEX_CLOUD_API_TOKEN

stocks = pd.read_csv('sp_500_stocks.csv')

output_to_console = False
output_to_excel = True
number_of_stocks = 20

# excel parameters
background_color = '#ffffff'
font_color = '#000000'
column_width_pixels = 15
excel_output_name = 'value_stocks'

# batch call
def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))


# calculating number of shares to buy
def portfolio_input():
    global portfolio_size
    portfolio_size = input("Enter the value of your portfolio:")

    try:
        val = float(portfolio_size)
    except ValueError:
        print("That's not a number! \n Try again:")
        portfolio_size = input("Enter the value of your portfolio:")


# dataframe using robust value
rv_columns = [
    'Ticker',
    'Company Name',
    #'Sector'
    'Price',
    'Num. Shares to Buy',
    'P/E Ratio',
    'P/E Percentile',
    'Forward P/E Ratio',
    'Forward P/E Percentile',
    'P/B Ratio',
    'P/B Percentile',
    'P/S Ratio',
    'P/S Percentile',
    'Debt/Equity Ratio',
    'D/E Percentile',
    'PEG',
    'PEG Percentile',
    'EV/EBITDA',
    'EV/EBITDA Percentile',
    'EV/GP',
    'EV/GP Percentile',
    'Put/Call Ratio',
    'Put/Call Percentile',
    'RV Score'
]

rv_dataframe = pd.DataFrame(columns = rv_columns)

for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote,company,advanced-stats&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    #print(data['AAPL']['advanced-stats']['putCallRatio'])
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
                    data[symbol]['quote']['companyName'],
                    #data[symbol]['company']['sector'],
                    data[symbol]['quote']['latestPrice'],
                    'N/A',
                    data[symbol]['quote']['peRatio'],
                    'N/A',
                    data[symbol]['advanced-stats']['forwardPERatio'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToBook'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToSales'],
                    'N/A',
                    data[symbol]['advanced-stats']['pegRatio'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToSales'],
                    'N/A',
                    ev_to_ebitda,
                    'N/A',
                    ev_to_gross_profit,
                    'N/A',
                    data[symbol]['advanced-stats']['putCallRatio'],
                    'N/A',
                    'N/A'
                ],
                index = rv_columns
            ),
            ignore_index = True
        )


# missing values

for column in ['P/E Ratio', 'Forward P/E Ratio','P/B Ratio', 'P/S Ratio', 'Debt/Equity Ratio', 'PEG', 'EV/EBITDA', 'EV/GP', 'Put/Call Ratio']:
    rv_dataframe[column].fillna(rv_dataframe[column].mean(), inplace = True)

# calculating value percentiles
from scipy.stats import percentileofscore as score
metrics = {'P/E Ratio': 'P/E Percentile',
            'Forward P/E Ratio': 'Forward P/E Percentile',
            'P/B Ratio': 'P/B Percentile',
            'P/S Ratio': 'P/S Percentile',
            'Debt/Equity Ratio': 'D/E Percentile',
            'PEG': 'PEG Percentile',
            'EV/EBITDA': 'EV/EBITDA Percentile',
            'EV/GP': 'EV/GP Percentile',
            'Put/Call Ratio': 'Put/Call Percentile'
           }

for metric in metrics.keys():
    for row in rv_dataframe.index:
        rv_dataframe.loc[row, metrics[metric]] = score(rv_dataframe[metric], rv_dataframe.loc[row, metric])/100

# calculating RV score
for row in rv_dataframe.index:
    value_percentiles = []
    for metric in metrics.keys():
        value_percentiles.append(rv_dataframe.loc[row, metrics[metric]])
    rv_dataframe.loc[row, 'RV Score'] = mean(value_percentiles)

# selecting 50 best value stocks
rv_dataframe.sort_values('RV Score', ascending = True, inplace = True)
rv_dataframe = rv_dataframe[rv_dataframe['P/E Ratio'] > 0]
rv_dataframe = rv_dataframe[rv_dataframe['P/B Ratio'] > 0]
rv_dataframe = rv_dataframe[rv_dataframe['P/E Ratio'] > rv_dataframe['Forward P/E Ratio']]
rv_dataframe = rv_dataframe[rv_dataframe['Forward P/E Ratio'] > 0]
rv_dataframe = rv_dataframe[rv_dataframe['Debt/Equity Ratio'] > 0]
rv_dataframe = rv_dataframe[:number_of_stocks]
rv_dataframe.reset_index(drop = True, inplace = True)

# calculating number of shares to buy
portfolio_input()
position_size = float(portfolio_size) / len(rv_dataframe.index)
for row in rv_dataframe.index:
    rv_dataframe.loc[row, 'Num. Shares to Buy'] = math.floor(position_size / rv_dataframe.loc[row, 'Price'])
if output_to_console:
    print(rv_dataframe)

# excel output
writer = pd.ExcelWriter(excel_output_name+'.xlsx', engine = 'xlsxwriter')
rv_dataframe.to_excel(writer, sheet_name = "Value Strategy", index = False)

string_template = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_template = writer.book.add_format(
        {
            'num_format': '$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_template = writer.book.add_format(
        {
            'num_format': '0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

float_template = writer.book.add_format(
        {
            'num_format': '0.0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

percent_template = writer.book.add_format(
        {
            'num_format': '0.0%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

column_formats = {
                    'A': ['Ticker', string_template],
                    'B': ['Company Name', string_template],
                   # 'C': ['Sector', string_template],
                    'C': ['Price', dollar_template],
                    'D': ['Num. Shares to Buy', integer_template],
                    'E': ['P/E Ratio', float_template],
                    'F': ['P/E Percentile', percent_template],
                    'G': ['Forward P/E Ratio', float_template],
                    'H': ['Forward P/E Percentile', percent_template],
                    'I': ['P/B Ratio', float_template],
                    'J': ['P/B Percentile', percent_template],
                    'K': ['P/S Ratio', float_template],
                    'L': ['P/S Percentile', percent_template],
                    'M': ['Debt/Equity Ratio', float_template],
                    'N': ['D/E Percentile', percent_template],
                    'O': ['PEG', float_template],
                    'P': ['PEG Percentile', percent_template],
                    'Q': ['EV/EBITDA', float_template],
                    'R': ['EV/EBITDA Percentile', percent_template],
                    'S': ['EV/GP', float_template],
                    'T': ['EV/GP Percentile', percent_template],
                    'U': ['Put/Call Ratio', float_template],
                    'V': ['Put/Call Percentile', percent_template],
                    'W': ['RV Score', percent_template]
                  }
for column in column_formats.keys():
    writer.sheets['Value Strategy'].set_column(f'{column}:{column}', column_width_pixels, column_formats[column][1])
    writer.sheets['Value Strategy'].write(f'{column}1', column_formats[column][0], column_formats[column][1])
if output_to_excel:
    writer.save()
