import pandas as pd
import yfinance as yf
import numpy as np
from datetime import datetime, date, timezone, timedelta
from McScreener import get_stats



firstdate = '2015-10-01'
lastdate = '2025-09-01'

SnPdf = yf.download(
    tickers='^GSPC',
    start=firstdate,
    end=lastdate,
    interval='1mo',
    group_by='Ticker',
    auto_adjust=False,
    progress=False
    )
print('Fresh download...')
print(SnPdf)
SnPdf = SnPdf.droplevel(level='Ticker',axis=1)
print('After droplevel...')
print(SnPdf)
SnPdf.columns = SnPdf.columns.str.replace(' ','_')
SnPdf.columns = SnPdf.columns.str.lower()
print('After column fix...')
print(SnPdf)
#stats = get_stats(SnPdf)
#print(stats)
