# import libraries
import numpy as np  # executes super fast because np is written in C
import pandas as pd
import xlsxwriter
import requests
import math
from secrets import IEX_CLOUD_API_TOKEN
""" 

A Equal-Weight S&P500 Index Fund instead of market-capitalization

A Python script that will accept the value of a portfolio and tell one how many shares of each S&P 500 constituent 
one should purchase to get an equal-weight version of the index fund.


# currently in sandbox mode bc free
# so all of the data is currently randomized and not representative of reality
# thus secrets file with API key is public
# if using API in real mode, make sure to keep the key secret
"""


stocks = pd.read_csv("sp_500_stocks.csv")


# symbol = "AAPL"
# api_url = f"https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}"
# data = requests.get(api_url).json() # make it easily accessible
# # build a pd dataframe with several columns
# price = data['latestPrice'] # all data here is randomized
# market_cap = data["marketCap"]

my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns=my_columns)

# batch api
# split lists into sublists of len 100
def chunks(lst, n):
    """Yild successive n sized chuncks from lst"""
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

symbol_groups = list(chunks(stocks['Ticker'], 100))  #list of pandas series
symbol_strings = []

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        # in data ['quote'] is required bc it is the API endpoint; we're looping thprugh API calls
        final_dataframe = final_dataframe.append(pd.Series([symbol, data[symbol]['quote']['latestPrice'], data[symbol]['quote']['marketCap'],'N/A'], index = my_columns),ignore_index=True)


val = 1000
position_size = val/(len(final_dataframe.index))
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number Of Shares to Buy'] = math.floor(position_size / final_dataframe['Stock Price'][i])

writer = pd.ExcelWriter('Recommended Trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index=False)

# color scheme for excel sheet, using html hex codes
background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        "border": 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        "border": 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        "border": 1
    }
)

writer.save()
column_formats = {
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)
writer.save()