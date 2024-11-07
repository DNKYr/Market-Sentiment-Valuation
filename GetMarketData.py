import yfinance as yf
import pandas as pd
import xlsxwriter

data = pd.read_csv('constituents.csv')
constituents = data.values

my_columns = ['Ticker', 'Price', '50 Day Average', '200 Day Average']
L = []
for stock in data['Symbol']:
    d = yf.Ticker(stock).info
    save_series = pd.Series([stock, d['regularMarketPreviousClose'], d['fiftyDayAverage'], d['twoHundredDayAverage']], index = my_columns)
    L.append(save_series)

final_dataframe = pd.DataFrame(L)

#formatting excel output
with pd.ExcelWriter('MarketData.xlsx') as writer:
    final_dataframe.to_excel(writer, sheet_name = 'Data', index = False)

    # Access the worksheet 'recommended trade'
    worksheet = writer.sheets['Data']