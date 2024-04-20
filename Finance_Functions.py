import pandas as pd

import config
from fredapi import Fred
fred = Fred(api_key=config.fredapi_key())

"""
Functions
"""





"""
APIs

"""

def fred_data(data_id, start=None, end=None, data_file=None):
    """
    fred API: http://erwan.marginalq.com/index_files/tea_files/FRED_API.html
    
    'data_id'
      The argument 'series_id' in the original function, fred.get_series()
    'start'
      observation start date; must be 'yyyy-mm-dd' format
    'end'
      observation end date; must be 'yyyy-mm-dd' format
      
    """
    series = fred.get_series(series_id = data_id, 
                             observation_start = start,
                             observation_end = end)
    
    if data_file != None:
        df = pd.read_csv(data_file, parse_dates=['Date'])
        col = list(df.columns)
        latest = pd.DataFrame({col[0]: series.index, col[1]: series.values})
        df = pd.concat([df, latest], ignore_index=True)
        df['Date'] = pd.to_datetime(df['Date'])
        df.to_csv(data_file, index=False)
        return df
        
    return series