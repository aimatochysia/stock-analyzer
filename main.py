import os
import time
import requests
import yfinance as yf
import pandas as pd
import math as m
import csv
from git import Repo
from dotenv import load_dotenv
from datetime import datetime
from io import StringIO

load_dotenv()
GITHUB_REPO = os.getenv('_GITHUB_REPO')
GITHUB_TOKEN = os.getenv('_GITHUB_TOKEN')
BRANCH_NAME = os.getenv('_BRANCH_NAME') 
TEMP_DIR = os.path.join(os.getcwd(), 'repo')
if not os.path.exists(TEMP_DIR):
    repo = Repo.clone_from(f'https://{GITHUB_TOKEN}@github.com/{GITHUB_REPO}.git', TEMP_DIR, branch=BRANCH_NAME)
else:
    repo = Repo(TEMP_DIR)

def load_stock_symbols(filename='stocklist.xlsx'):
    df = pd.read_excel(filename, usecols=[1])
    stock_symbols = df.iloc[1:].dropna().squeeze().tolist()
    return stock_symbols

all_values = load_stock_symbols()
stocks = [item + ".JK" for item in all_values]

def fetch_stock_data(stock):
    for period in ['max', '1y', '1mo']:
        try:
            history = yf.download(stock, period=period)
            if not history.empty:
                return history
        except Exception as e:
            continue
    raise ValueError(f"Failed to fetch data for {stock}.")

def get_perc_change(n_days, history_cls):
    temp_array = history_cls[:n_days]
    first_item = temp_array[0]
    last_item = temp_array[-1]
    perc_change = (first_item - last_item) / last_item * 100
    return perc_change

def get_volas(n_days, history_cls):
    temp_price = history_cls[:n_days]
    avg_price = sum(temp_price) / n_days
    diff_cls = [temp_price[i] - avg_price for i in range(n_days)]
    diff_cls = [diff**2 for diff in diff_cls]
    variance = sum(diff_cls) / n_days
    std_dev = m.sqrt(variance)
    return std_dev

def get_avg_vol(n_days, history_vol):
    return sum(history_vol[:n_days]) / n_days

def get_ma(n_days, history_cls):
    return sum(history_cls[:n_days]) / n_days

def calculate_display_ma(ma5, ma10, ma20, ma50, ma100, ma200, stock):
    symbols = []
    output_const = 0
    for ma in [ma5, ma10, ma20, ma50, ma100, ma200]:
        if stock > ma:
            symbols.append(">")
            output_const += 1
        elif stock == ma:
            symbols.append("=")
        else:
            symbols.append("<")
            output_const -= 1
    return symbols, output_const

def analyze_bound_stock(n_days, history_cls, history_vol):
    output_const = 0.0
    i = 0
    is_avg_check = False
    multiplier = n_days / 2
    temp_price = history_cls[:n_days]
    temp_vol = history_vol[:n_days]
    avg_price = sum(temp_price) / n_days
    avg_vol = sum(temp_vol) / n_days
    upper_bound_avg_vol = avg_vol + avg_vol / multiplier
    lower_bound_avg_vol = avg_vol - avg_vol / multiplier
    lower_bound = avg_price - (avg_price / multiplier)
    while not is_avg_check and i < n_days:
        is_low_vol = lower_bound_avg_vol < history_vol[i] < upper_bound_avg_vol
        if lower_bound < history_cls[i+1] and is_low_vol:
            i += 1
            output_const += 1
        else:
            is_avg_check = True
    return output_const

current_date = datetime.now().strftime("%d-%m-%Y")
csv_filename = f"stock_data_{current_date}.csv"
debug_filename = f"debug_stock_scrapper_{current_date}.txt"
csv_file = StringIO()
debug_file = StringIO()

header = [
    'Stock', 'Market Cap and Buy Analysis', 'Buy Analysis', 'Volume Analysis Result',
    'Volume Analysis 5 Days', 'Volume Analysis 20 Days', 'Volume Analysis 50 Days',
    'MA Analysis', 'MA5', 'MA5 Symbol', 'MA10', 'MA10 Symbol', 'MA20', 'MA20 Symbol',
    'MA50', 'MA50 Symbol', 'MA100', 'MA100 Symbol', 'MA200', 'MA200 Symbol',
    '3D Change%', '5D Change%', '20D Change%', 'Yesterday Closing Price', 'Current Price',
    'Volatility in 3 Day', 'Volatility in 5 Days', 'Volatility in 20 Days',
    'Market Cap Value', 'Volume Average in 5 Days', 'Volume Average in 20 Days',
    'Volume Average in 50 Days', 'Volume Average in 100 Days',
]

csv_writer = csv.DictWriter(csv_file, fieldnames=header)
csv_writer.writeheader()

for stock in stocks:
    print(f"Currently processing: {stock}")
    try:
        history = fetch_stock_data(stock)
        stock_data = yf.Ticker(stock)
        info = stock_data.info
        history_cls = history['Close'].tolist()[::-1]
        history_vol = history['Volume'].tolist()[::-1]

        perc3 = get_perc_change(3, history_cls)
        perc5 = get_perc_change(5, history_cls)
        perc20 = get_perc_change(20, history_cls)

        vola3 = get_volas(3, history_cls)
        vola5 = get_volas(5, history_cls)
        vola20 = get_volas(20, history_cls)

        vol5 = get_avg_vol(5, history_vol)
        vol20 = get_avg_vol(20, history_vol)
        vol50 = get_avg_vol(50, history_vol)
        vol100 = get_avg_vol(100, history_vol)

        ma5 = get_ma(5, history_cls)
        ma10 = get_ma(10, history_cls)
        ma20 = get_ma(20, history_cls)
        ma50 = get_ma(50, history_cls)
        ma100 = get_ma(100, history_cls)
        ma200 = get_ma(200, history_cls)

        display_ma, ma_const = calculate_display_ma(ma5, ma10, ma20, ma50, ma100, ma200, history_cls[0])

        volume_const5 = analyze_bound_stock(5, history_cls, history_vol)
        volume_const20 = analyze_bound_stock(20, history_cls, history_vol)
        volume_const50 = analyze_bound_stock(50, history_cls, history_vol)

        market_cap = info.get('marketCap', 'N/A')
        if market_cap != 'N/A':
            if isinstance(market_cap, str) and market_cap.endswith('B'):
                market_cap = float(market_cap[:-1]) * 1e9
            elif isinstance(market_cap, str) and market_cap.endswith('T'):
                market_cap = float(market_cap[:-1]) * 1e12
            else:
                market_cap = float(market_cap)

        data = {
            'Stock': stock,
            'Market Cap and Buy Analysis': market_cap,
            'Buy Analysis': ma_const,
            'Volume Analysis Result': volume_const50,
            'Volume Analysis 5 Days': volume_const5,
            'Volume Analysis 20 Days': volume_const20,
            'Volume Analysis 50 Days': volume_const50,
            'MA Analysis': display_ma,
            'MA5': ma5,
            'MA5 Symbol': display_ma[0],
            'MA10': ma10,
            'MA10 Symbol': display_ma[1],
            'MA20': ma20,
            'MA20 Symbol': display_ma[2],
            'MA50': ma50,
            'MA50 Symbol': display_ma[3],
            'MA100': ma100,
            'MA100 Symbol': display_ma[4],
            'MA200': ma200,
            'MA200 Symbol': display_ma[5],
            '3D Change%': perc3,
            '5D Change%': perc5,
            '20D Change%': perc20,
            'Yesterday Closing Price': history_cls[1],
            'Current Price': history_cls[0],
            'Volatility in 3 Day': vola3,
            'Volatility in 5 Days': vola5,
            'Volatility in 20 Days': vola20,
            'Market Cap Value': market_cap,
            'Volume Average in 5 Days': vol5,
            'Volume Average in 20 Days': vol20,
            'Volume Average in 50 Days': vol50,
            'Volume Average in 100 Days': vol100,
        }

        csv_writer.writerow(data)
    except Exception as e:
        debug_file.write(f"Error processing {stock}: {str(e)}\n")

def push_to_github(filename, content):
    file_path = os.path.join(TEMP_DIR, filename)
    with open(file_path, 'w') as file:
        file.write(content.getvalue())

    repo.index.add([file_path])
    repo.index.commit(f'Update {filename}')
    origin = repo.remote(name='origin')
    origin.push()

push_to_github(csv_filename, csv_file)
push_to_github(debug_filename, debug_file)
