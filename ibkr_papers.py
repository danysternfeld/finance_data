from  foreign_papers import *
from pandas import DataFrame
import pandas as pd


def get_ibkr_sharpe_and_performance():
    """
    This function reads the IBKR holdings and ticker mapping from CSV files,
    computes the Sharpe ratio and performance for each paper, and saves the
    enriched data to a new CSV file.
    """
    ibkr_file = "Open_Positions.csv"
    ticker_file = "TICKER in YF.xlsx"
    tickers_df = pd.read_excel(ticker_file)
    ibkr_df = pd.read_csv(ibkr_file)
    today = pd.Timestamp.now().strftime("%Y-%m-%d")
    last_year = (pd.Timestamp.now() - pd.DateOffset(years=1)).strftime("%Y-%m-%d")
    for paper in ibkr_df['Symbol']:
        if paper in tickers_df['TICKER'].values:
            yf_ticker = str(tickers_df.loc[tickers_df['TICKER'] == paper]["YF TICKER"].values[0])
            
            sharpe_ratio = compute_sharpe_ratio(yf_ticker,  last_year,today)
            performance = compute_performance(yf_ticker)
            ibkr_df.loc[ibkr_df['Symbol'] == paper, 'sharpe_ratio_12_months'] = sharpe_ratio
            ibkr_df.loc[ibkr_df['Symbol'] == paper, 'performance_1y'] = performance.get("1y", None)
            ibkr_df.loc[ibkr_df['Symbol'] == paper, 'performance_3y'] = performance.get("3y", None)
    ibkr_df.to_excel("IBKR_HOLDINGS.xlsx",engine="openpyxl")


if __name__ == "__main__":
    get_ibkr_sharpe_and_performance()


