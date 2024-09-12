import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

# Define the list of ticker symbols
tickers = ["SPY", "QQQ", "VOO", "DIA", "IWM"]

# Define today's date as the start date
start_date = datetime.today()  # Start date

# Define the end date as start_date + 365 days
end_date = start_date + timedelta(days=365)  # End date: 365 days from start_date

# Create empty lists to store calls and puts dataframes
all_calls_list = []
all_puts_list = []

# Iterate over the tickers
for ticker in tickers:
    # Create a yfinance ticker object
    stock = yf.Ticker(ticker)

    # Create empty lists to store calls and puts dataframes for the current ticker
    calls_list = []
    puts_list = []

    # Iterate over the range of dates
    current_date = start_date
    while current_date <= end_date:
        expiration_date = current_date.strftime("%Y-%m-%d")

        try:
            # Get the options data for the current expiration date
            options = stock.option_chain(expiration_date)

            # Access the calls and puts dataframes
            calls = options.calls
            puts = options.puts

            # Add "Expiration Date" column to calls and puts dataframes
            calls['Expiration Date'] = expiration_date
            puts['Expiration Date'] = expiration_date

            # Append calls and puts dataframes to the respective lists for the current ticker
            calls_list.append(calls)
            puts_list.append(puts)
        except ValueError:
            print(f"Expiration {expiration_date} cannot be found for ticker {ticker}. Skipping to the next date.")

        # Move to the next date
        current_date += timedelta(days=1)

    # Concatenate all calls and puts dataframes for the current ticker into a single dataframe
    all_calls = pd.concat(calls_list)
    all_puts = pd.concat(puts_list)

    # Append the calls and puts dataframes for the current ticker to the respective lists for all tickers
    all_calls_list.append(all_calls)
    all_puts_list.append(all_puts)

# Concatenate all calls and puts dataframes for all tickers into a single dataframe
all_calls_df = pd.concat(all_calls_list)
all_puts_df = pd.concat(all_puts_list)

# Print the options data
print("All Calls:")
print(all_calls_df)

print("\nAll Puts:")
print(all_puts_df)

all_calls_df['lastTradeDate'] = all_calls_df['lastTradeDate'].dt.tz_convert(None)
all_calls_df.to_excel(r"C:\Users\MPM II\OneDrive\Desktop\Financial Equation\index_calls.xlsx", index=False)

all_puts_df['lastTradeDate'] = all_puts_df['lastTradeDate'].dt.tz_convert(None)
all_puts_df.to_excel(r"C:\Users\MPM II\OneDrive\Desktop\Financial Equation\index_puts.xlsx", index=False)

import pandas as pd
import yfinance as yf

# Define the list of tickers
tickers = ["SPY", "QQQ", "VOO", "DIA", "IWM", "^IRX"]

# Create a yfinance ticker object for each ticker
ticker_objects = [yf.Ticker(ticker) for ticker in tickers]

# Get the historical data for each ticker
data = [ticker.history(period="1d")["Close"] for ticker in ticker_objects]

# Combine the data into a DataFrame
df = pd.concat(data, axis=1)
df.columns = tickers

# Reset the index and convert the columns to a MultiIndex
df = df.reset_index().melt(id_vars=["Date"], var_name="Ticker", value_name="Price")
df = df[["Ticker", "Price"]]

# Print the DataFrame
print(df)

import pandas as pd
import datetime

# Specify the file path
file_path = r"C:\Users\MPM II\OneDrive\Desktop\Financial Equation\index_calls.xlsx"

# Read the Excel file
data = pd.read_excel(file_path)

# Convert the "Expiration" and reference date columns to datetime
data["Expiration Date"] = pd.to_datetime(data["Expiration Date"])
# Get today's date
reference_date = pd.Timestamp.now().normalize()

# Calculate the time to maturity
data["time_to_maturity"] = ((data["Expiration Date"] - reference_date).dt.days + 1)/365

# Extract the Ticker from the 'contractSymbol' column
data['Tickers_Symbol'] = data['contractSymbol'].str.slice(stop=3)

# Create a new DataFrame with the prices matched based on the Ticker
merged_data = pd.merge(data, df, left_on="Tickers_Symbol", right_on="Ticker", how="left")

# Create moneyness column
merged_data['moneyness'] = merged_data['strike']/merged_data['Price']

# Add ^IRX price as a new column named "Rf_Rate" to the merged_data dataframe
irx_price = df[df['Ticker'] == '^IRX']['Price'].iloc[0]
merged_data['Rf_Rate'] = irx_price

# Print the updated DataFrame
print(merged_data)

# Get today's date
today = datetime.date.today()

# Format the date as "YYYY-MM-DD"
date_string = today.strftime("%Y-%m-%d")

# Create the file path with today's date
file_path = fr"C:\Users\MPM II\OneDrive\Desktop\Financial Equation\call_option_inputs_{date_string}.xlsx"

# Export merged_data to Excel using the updated file path
merged_data.to_excel(file_path, index=False)
