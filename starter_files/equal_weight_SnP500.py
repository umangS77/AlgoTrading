import numpy as numpy
import pandas as pd
import math
import requests
import xlsxwriter

stocks = pd.read_csv('./sp_500_stocks.csv')
stocks = stocks[~stocks['Ticker'].isin(['DISCA', 'HFC','VIAC','WLTW'])]

from secrets import IEX_CLOUD_API_TOKEN

my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns = my_columns)

# for symbol in stocks['Ticker']:
# 	api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
# 	data = requests.get(api_url).json()
# 	price = data['latestPrice']
# 	market_cap = data['marketCap']/1000000000
# 	final_dataframe =  final_dataframe.append(
# 		pd.Series(
# 		[
# 			symbol,
# 			price,
# 			market_cap,
# 			'N/A'
# 		],
# 		index = my_columns,
# 		),
# 		ignore_index = True
# 	)

# print(final_dataframe)


####################################################
# batch API calls

def chunks(lst, n):
	# yield successive n sized chunks
	for i in range(0,len(lst),n):
		yield lst[i:i+n]


symbol_groups = list(chunks(stocks['Ticker'], 100)) # list of lists of 100 stocks
symbol_strings = []
for i in range(0,len(symbol_groups)):
	symbol_strings.append(','.join(symbol_groups[i]))


for symbol_string in symbol_strings:
	batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
	# print(batch_api_call_url)
	data = requests.get(batch_api_call_url).json()
	for symbol in symbol_string.split(','):
		# print(symbol)
		# price = data[symbol]['quote']['latestPrice']
		# market_cap = data[symbol]['quote']['marketCap']
		final_dataframe =  final_dataframe.append(
			pd.Series(
			[
				symbol,
				data[symbol]['quote']['latestPrice'],
				data[symbol]['quote']['marketCap'],
				'N/A'
			],
			index = my_columns,
			),
			ignore_index = True
		)


# print(final_dataframe)

##############################################################################


portfolio_size = input('Enter the value of your portfolio: ')

try:
	val = float(portfolio_size)
except ValueError:
	print('Please Enter Numeric Value')
	portfolio_size = input('Enter the value of your portfolio: ')
	val = float(portfolio_size)


position_size = val/len(final_dataframe.index)
# print(position_size)

for i in range(0,len(final_dataframe.index)):
	final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])


# print(final_dataframe)

writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
	{
		'font_color' : font_color,
		'bg_color' : background_color,
		'border' : 1
	}
)

dollar_format = writer.book.add_format(
	{
		'num_format' : '$ 0.00',
		'font_color' : font_color,
		'bg_color' : background_color,
		'border' : 1
	}
)

integer_format = writer.book.add_format(
	{
		'num_format' : '0',
		'font_color' : font_color,
		'bg_color' : background_color,
		'border' : 1
	}
)



# writer.sheets['Recommended Trades'].set_column('A:A', 10, string_format)
# writer.sheets['Recommended Trades'].set_column('B:B', 18, dollar_format)
# writer.sheets['Recommended Trades'].set_column('C:C', 32, dollar_format)
# writer.sheets['Recommended Trades'].set_column('D:D', 32, integer_format)

# writer.sheets['Recommended Trades'].write('A1', 'Ticker', string_format)
# writer.sheets['Recommended Trades'].write('B1', 'Stock Price', dollar_format)
# writer.sheets['Recommended Trades'].write('C1', 'Market Capitalization', dollar_format)
# writer.sheets['Recommended Trades'].write('D1', 'Number of Shares to Buy', integer_format)


column_formats = {
	'A':['Ticker', string_format],
	'B':['Stock Price', dollar_format],
	'C':['Market Capitalization', dollar_format],
	'D':['Number of Shares to Buy', integer_format]
}

for col in column_formats.keys():
	writer.sheets['Recommended Trades'].set_column(f'{col}:{col}', 32, column_formats[col][1])
	writer.sheets['Recommended Trades'].write(f'{col}1', column_formats[col][0], column_formats[col][1])

writer.save()