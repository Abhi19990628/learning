# Required Libraries
import pandas as pd
import numpy as np
import yfinance as yf
import talib
import matplotlib.pyplot as plt

# Step 1: Scrape EUR/INR currency data from Yahoo Finance
start_date = '2023-01-01'
end_date = '2024-02-16'
currency_data = yf.download('EURINR=X', start=start_date, end=end_date)

# Step 2: Technical Analysis
# Calculate Moving Average
currency_data['MA_20'] = talib.SMA(currency_data['Close'], timeperiod=20)
currency_data['MA_50'] = talib.SMA(currency_data['Close'], timeperiod=50)

# Calculate Bollinger Bands
currency_data['upper_band'], currency_data['middle_band'], currency_data['lower_band'] = talib.BBANDS(
    currency_data['Close'], timeperiod=20)

# Calculate Commodity Channel Index (CCI)
currency_data['CCI'] = talib.CCI(currency_data['High'], currency_data['Low'], currency_data['Close'], timeperiod=20)

# Step 3: Decision Making
# Assuming trading strategy: Buy when price > 20-day MA, Sell when price < 20-day MA, Neutral otherwise
currency_data['Buy_Sell_20MA'] = np.where(currency_data['Close'] > currency_data['MA_20'], 'BUY', 'SELL')
currency_data['Buy_Sell_50MA'] = np.where(currency_data['Close'] > currency_data['MA_50'], 'BUY', 'SELL')

# Assuming trading strategy: Buy when price < lower Bollinger Band, Sell when price > upper Bollinger Band, Neutral otherwise
currency_data['Buy_Sell_Bollinger'] = np.where(currency_data['Close'] < currency_data['lower_band'], 'BUY',
                                                np.where(currency_data['Close'] > currency_data['upper_band'], 'SELL', 'NEUTRAL'))

# Assuming trading strategy: Buy when CCI > 100, Sell when CCI < -100, Neutral otherwise
currency_data['Buy_Sell_CCI'] = np.where(currency_data['CCI'] > 100, 'BUY',
                                         np.where(currency_data['CCI'] < -100, 'SELL', 'NEUTRAL'))

# Filter data for analysis on February 16, 2024
analysis_date = '2024-02-16'
analysis_data = currency_data.loc[analysis_date]

# Display results
print("Results for February 16, 2024:")
print("Moving Average (20-day):", analysis_data['Buy_Sell_20MA'])
print("Moving Average (50-day):", analysis_data['Buy_Sell_50MA'])
print("Bollinger Bands:", analysis_data['Buy_Sell_Bollinger'])
print("CCI:", analysis_data['Buy_Sell_CCI'])

# Visualization
# Plotting Moving Averages
plt.figure(figsize=(12, 6))
plt.plot(currency_data.index, currency_data['Close'], label='Close Price', color='blue')
plt.plot(currency_data.index, currency_data['MA_20'], label='20-day MA', color='red')
plt.plot(currency_data.index, currency_data['MA_50'], label='50-day MA', color='green')
plt.title('EUR/INR Close Price and Moving Averages')
plt.xlabel('Date')
plt.ylabel('Price')
plt.legend()
plt.show()

# Plotting Bollinger Bands
plt.figure(figsize=(12, 6))
plt.plot(currency_data.index, currency_data['Close'], label='Close Price', color='blue')
plt.plot(currency_data.index, currency_data['upper_band'], label='Upper Bollinger Band', color='red')
plt.plot(currency_data.index, currency_data['middle_band'], label='Middle Bollinger Band', color='green')
plt.plot(currency_data.index, currency_data['lower_band'], label='Lower Bollinger Band', color='orange')
plt.title('EUR/INR Close Price and Bollinger Bands')
plt.xlabel('Date')
plt.ylabel('Price')
plt.legend()
plt.show()

# Save results to PowerPoint
# You can export the relevant data and visualizations to PowerPoint using libraries like python-pptx.
import yfinance as yf

# Define the ticker symbol for EUR/INR
ticker_symbol = 'EURINR=X'

# Scrape currency data from Yahoo Finance
currency_data = yf.download(ticker_symbol, start='2023-01-01', end='2024-02-16')

# Display the first few rows of the data
print(currency_data.head())

