from email.errors import FirstHeaderLineIsContinuationDefect

from cffi import FFIError
from more_itertools import last
from mpl_toolkits.axes_grid1 import host_axes
import numpy as np
from sqlalchemy import true
import yfinance as yf
from pandas import DataFrame
import pandas as pd
import datetime
import sys
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)


def compute_sharpe_ratio(ticker: str, start_date: str, end_date: str, risk_free_rate: float = 0.055) -> float:
    """
    Compute the Sharpe ratio for a given stock ticker over a specified date range.

    Parameters:
    - ticker: Stock ticker symbol (e.g., 'AAPL').
    - start_date: Start date for historical data in 'YYYY-MM-DD' format.
    - end_date: End date for historical data in 'YYYY-MM-DD' format.
    - risk_free_rate: Annual risk-free rate (default is 5.5%).

    Returns:
    - Sharpe ratio as a float.
    """
    # Download historical data
    df = yf.download(ticker, start=start_date, end=end_date, progress=False)

    # Extract close prices and compute daily returns
    prices = df["Close"]
    returns = prices.pct_change(fill_method=None).dropna()

    # Convert annual risk-free rate to daily
    risk_free_daily = risk_free_rate / 252

    # Compute excess returns and annualize Sharpe ratio
    excess_returns = returns - risk_free_daily
    sharpe_ratio = np.sqrt(252) * (excess_returns.mean() / excess_returns.std())
    return sharpe_ratio.iloc[0]

def check_3year_perf_sanity(history_data):
    today = pd.Timestamp.now().strftime("%Y-%m-%d")
    three_years_ago = (pd.Timestamp.now() - pd.DateOffset(years=3)).strftime("%Y-%m")
    first_date = history_data.index[0].strftime("%Y-%m")
    return first_date == three_years_ago



def compute_performance(ticker: str) :
    periods = ["1y", "3y"]
    performance = {}
    sane = true
    #print(f"Ticker: {ticker}")
    for period in periods:
        data = yf.Ticker(ticker).history(period=period).dropna()
        if(period == "3y"):
            sane = check_3year_perf_sanity(data)
        if sane:
            start_price = data["Close"].iloc[0]
            end_price = data["Close"].iloc[-1]
            pct_change = (end_price - start_price) / start_price * 100
            performance[period] = pct_change
        else:
            print("\tMissing 3y data. will mark as NA")
            performance[period] = "NA"
        #print(performance)
    return performance



