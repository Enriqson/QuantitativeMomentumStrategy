import numpy as np
import pandas as pd
import requests
import math
from scipy.stats import percentileofscore as score
import xlsxwriter
from statistics import mean

stocks = pd.read_csv('sp_500_stocks.csv')
from secrets import IEX_CLOUD_API_TOKEN

#Configure and fetch Data

columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy',
    'Momentum Score',
    '1 year Return',
    '1 year Return Percentile',
    '6 month Return',
    '6 month Return Percentile',
    '3 month Return',
    '3 month Return Percentile',
    '1 month Return',
    '1 month Return Percentile'
]

final_dataframe = pd.DataFrame(columns=columns)

def divList(lst,n):
    for i in range(0,len(lst),n):
        yield lst[i:i+n]

symbol_groups = list(divList(stocks['Ticker'],100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))


for symbol_string in symbol_strings:

    batch_api_url=f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=price,stats&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_url).json()

    for symbol in symbol_string.split(","):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                symbol,
                data[symbol]['price'],
                'N/A',
                'N/A',
                data[symbol]['stats']['year1ChangePercent'],
                'N/A',
                data[symbol]['stats']['month6ChangePercent'],
                'N/A',
                data[symbol]['stats']['month3ChangePercent'],
                'N/A',
                data[symbol]['stats']['month1ChangePercent'],
                'N/A'
            ],
            index = columns
            ),
            ignore_index=True
        )


#Calculate Momentum Scores
time_periods = [
    '1 year',
    '6 month',
    '3 month',
    '1 month'
]

for time_period in time_periods:
    abs_col = f'{time_period} Return'
    percent_col = f'{time_period} Return Percentile'
    final_dataframe[abs_col] = final_dataframe[abs_col].fillna(0.0)
    final_dataframe[percent_col] = final_dataframe[abs_col].apply(lambda x:score(final_dataframe[abs_col],x))

for row in final_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(final_dataframe.loc[row,f'{time_period} Return Percentile'])
    final_dataframe.loc[row,'Momentum Score'] = mean(momentum_percentiles)

#Get top 50 highest scores
final_dataframe.sort_values('Momentum Score',ascending=False,inplace=True)
final_dataframe= final_dataframe[:50]
final_dataframe.reset_index(inplace=True,drop=True)


def portfolio_input():
    global portfolio_size
    portfolio_size = input('Enter the size of your portfolio: ')

    try:
        float(portfolio_size)
    except ValueError:
        print('Please enter a number!')
        portfolio_size = input('Enter the size of your portfolio: ')

portfolio_input()

portfolio_size = float(portfolio_size)

momentum_sum = final_dataframe['Momentum Score'].sum()

final_dataframe['Number of Shares to Buy'] = (final_dataframe['Momentum Score']*portfolio_size)/(final_dataframe['Price']*momentum_sum)
final_dataframe['Number of Shares to Buy'].apply(math.floor)

#Save to XLSX
writer = pd.ExcelWriter('recommended trades.xlsx', engine='xlsxwriter')

final_dataframe.to_excel(writer,'Recommended Trades',index=False)

background_color='#0a0a23'
font_color='#ffffff'

string_format = writer.book.add_format(
    {
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format':'$0.00',
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format':'0',
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

float_format = writer.book.add_format(
    {
        'num_format':'0.0000',
        'font_color':font_color,
        'bg_color':background_color,
        'border': 1
    }
)

column_formats={
    'A':['Ticker',string_format],
    'B':['Price',dollar_format],
    'C':['Number of Shares to Buy',integer_format],
    'D':['Momentum Score',float_format],
    'E':['1 year Return',float_format],
    'F':['1 year Return Percentile',float_format],
    'G':['6 month Return',float_format],
    'H':['6 month Return Percentile',float_format],
    'I':['3 month Return',float_format],
    'J':['3 month Return Percentile',float_format],
    'K':['1 month Return',float_format],
    'L':['1 month Return Percentile',float_format],
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}',18,column_formats[column][1])

writer.save()

print("End")