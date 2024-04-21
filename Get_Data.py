import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import datetime
from dateutil.relativedelta import relativedelta
import yfinance as yf
import os
import time


from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By
options = Options()
options.add_argument('--disable-extensions');
options.add_argument('--proxy-server="direct://"');
options.add_argument('--proxy-bypass-list=*');
options.add_argument('--start-maximized');

"""
Functions to retrieve data  
"""

def get_yahoo_finance_data(file_name:str, tickers:list, daily_change: str, volumes:list, columns:list):
    """
    (str, list, str, list, list) -> edit excel file
    Extract the data from yahoo finance and add them into excel file
    """
    file_list = os.listdir('data')
    if file_name not in file_list:
        return f'Error. There is no {file_name} in data'
    
    df = pd.DataFrame()
    file_path = f'data/{file_name}'
    book = xw.Book(fullname=file_path)
    for i in range(len(tickers)):
        sheet = book.sheets[tickers[i]]
        last_row = int(sheet.range('A1').end('down').row)
        last_date = sheet.range(f'A{last_row}').value
        if pd.Timestamp(last_date) == pd.Timestamp(datetime.date.today()):
            sheet.range(f'{last_row}:{last_row}').clear()
            last_row = int(sheet.range('A1').end('down').row)

        start_date = sheet.range(f'A{last_row}').value + datetime.timedelta(days=1)
        new_data = yf.download(tickers[i], start=start_date).reset_index()[columns] # Data:'Open, High, Low, Close'
        
        if daily_change != None:
            close_data = list(pd.DataFrame([sheet.range(f'{daily_change}{last_row}').value]).values)
            close_data += list(new_data[['Close']].values)

            close_func = lambda x: ((x.iloc[1]-x.iloc[0])/x.iloc[0])*100
            new_data['Change%'] = pd.Series(close_data).rolling(window=2).apply(close_func).round(2).dropna().values # Data:'Change%'
    
        if volumes != None:
            new_volume = yf.download(volumes[i], start=start_date).reset_index()[['Date', 'Volume']]
            new_data = pd.merge(new_data, new_volume, on='Date', how='left') # Data: 'Volume'
            
        cell_num = last_row+1
        for each_data in new_data.values:
            sheet.range(f'A{cell_num}').value = each_data # Colume 'Open', 'High', 'Low', 'Close', 'Change%', 'Volume'
            cell_num += 1
        
        new_data['Ticker'] = tickers[i] 
        df = pd.concat([df, new_data])
            
    book.save()
    print('Added data:')
    return df.reset_index().set_index(['Ticker', 'Date']).drop(['index'], axis=1)


def get_vix_futures(show_table = False, show_graph = False): 
    """
    Retrieve the different contracts' VIX futures from the Barchart.com
    """
    # Get original data
    text = web_scrape(url='https://www.barchart.com/futures/quotes/VIY00/futures-prices',
                      css_selector='#main-content-column > div > div.barchart-content-block.invisible.border-top-0.visible')
    raw_data = text.split('\n') # split original data as 'new line'
    columns = raw_data[:10]     # columns are the first 10 items 
    del raw_data[:11]           # remove unnecessary items
    data_list = []              # define list for each row data
    while len(raw_data) != 0:
        data_list.append(raw_data[:10])
        del raw_data[:10]        
    df = pd.DataFrame(data=data_list, columns=columns)
    
    # Adding new data (closing value) to the excel file
    book = xw.Book(fullname='data/Contrarian_Indicators.xlsx')
    sheet = book.sheets['VIX_Futures']
    row = int(sheet.range('A1').end('down').row)+1    # get the row for the new data
    df_for_date = pd.read_excel('data/Stock_Indices.xlsx') 
    date = [df_for_date['Date'].dropna().iloc[-1].date()] # get the date of the retrieved data
    values = list(df['Last'].apply(lambda x: x[:-1]).astype('float')) # convert values to float
    sheet.range(f'A{row}').value = date + values
    book.save()
    
    if show_graph == True:
        new_df = pd.read_excel('data/Contrarian_Indicators.xlsx', sheet_name='VIX_Futures').T
        new_df.rename(columns=new_df.iloc[0], inplace=True) # assign date as eacn column name
        new_df.drop('Date', inplace=True) # drop 'Date' row
        plt.plot(new_df.iloc[:, -5:]) # showing the latest 10 data
        plt.legend(list(new_df.columns.astype('str'))[-5:])
            
    if show_table == True:
        return df

def get_sp500_sector():
    sectors = {'CONS_DESC': '^SP500-25', 'CONS_STPL': '^SP500-30', 'ENERGY': '^SP500-1010',
               'FINANCIALS': '^SP500-40', 'HEALTH': '^SP500-35', 'INDUSTRIALS': '^SP500-20',
               'MATERIALS': '^SP500-15', 'REAL_ESTATE': '^SP500-60', 'TECHNOLOGY': '^SP500-45',
               'TELECOM_SVS': '^SP500-50', 'UTILITIES': '^SP500-55'}
    book = xw.Book('data/Stock_Indices.xlsx')
    sheet = book.sheets['S&P500_SECTOR']
    last_row = int(sheet.range('A1').end('down').row)
    last_day = sheet.range(f'A{last_row}').value.date()
    if datetime.date.today() == last_day:
        sheet.range(f'{last_row}:{last_row}').clear()
        last_row = last_row - 1 
    
    start_day = sheet.range(f'A{last_row}').value.date() + datetime.timedelta(days=1)
    prices = [yf.download(sectors[sector], start=start_day)["Close"].rename(sector) for sector in sectors]
    df_sector = pd.concat(prices, axis=1).reset_index().round(2)
    sheet.range(f'A{last_row+1}').value = df_sector.values
    book.save()
    print('Added data')
    return df_sector
            
            
def get_pcr():
    """
    Extract Total, Index, and Equity Put/Call Ratio
    Reference -> https://www.cboe.com/us/options/market_statistics/daily/        
    """
    book = xw.Book(fullname='data/Contrarian_Indicators.xlsx')
    sheet = book.sheets['PCR']
    last_row = int(sheet.range('A1').end('down').row)
    last_day = sheet.range(f'A{last_row}').value
    
    add_df = pd.DataFrame(columns=['Date','TOTAL','INDEX','EQUITY']) 
    sample = pd.read_excel('data/Stock_Indices.xlsx', sheet_name="^GSPC") # load 'sample' to get the days the market was open
    
    for day in sample[sample['Date'] > pd.Timestamp(last_day)]['Date']: # day: market dates that aren't  written in PCR_sheet
        data = web_scrape(url=f'https://www.cboe.com/us/options/market_statistics/daily/?dt={str(day.date())}',
                          css_selector='#daily-market-statistics > div > div:nth-child(2) > table')
        # only three items whose first item is 'TOTAL' or 'INDEX' or 'EQUITY'
        data_list = list(filter(lambda x: x.split(' ')[0] in ['TOTAL','INDEX','EQUITY'], data.split('\n')))        
        add_dict = {'Date': day.date()}
        for d in data_list:
            add_dict[d.split(' ')[0]] = d.split(' ')[-1]
            
        add_df.loc[len(add_df)] = add_dict
    
    # add option data into excel sheet
    sheet.range(f'A{last_row+1}').value = add_df.values
    book.save()
    print("Added data:")
    return add_df

    
def get_options(tickers: list):
    """
    Ticker examples:
    S&P500:     '$SPX'
    Nasdaq100:  '$IUXX'
    Dow Jones:  '$DJX'
    """
    book = xw.Book(fullname='data/Options.xlsx')
    day = [datetime.datetime.today().date()]
    
    for ticker in tickers:        
        text = web_scrape(url=f'https://www.barchart.com/stocks/quotes/{ticker}',
                             css_selector='#main-content-column > div > div.barchart-content-block.symbol-fundamentals.bc-cot-table-wrapper > div.block-content')
        data = text.split('\n')        
        values = []
        for v in [data[i] for i in range(1, len(data), 2)]:
            p = v.find('%')
            value = v[:p] if p != -1 else v
            values.append(value)
        if ticker == tickers[0]:
            columns = [data[i] for i in range(0, len(data), 2)]
            df = pd.DataFrame(columns=columns)
        df.loc[ticker] = values

        sheet = book.sheets[ticker]
        row = int(sheet.range('A1').end('down').row)+1
        sheet.range(f'A{row}').value = day + values
    
    book.save()
    print('Added data:')
    return df

    
def get_aaii(load_sheet_delete=True):
    """
    Downloaded excel file before running this function will be removed, if its input is still True.
    Reference -> https://www.aaii.com/sentimentsurvey
    """
    load_sheet = pd.read_excel(io='C:/Users/runru/Downloads/sentiment.xls', sheet_name='SENTIMENT').iloc[4:, :4]
    load_sheet.columns = ['Date', 'Bullish', 'Neutral', 'Bearish']
    load_sheet_lday = list(filter(lambda x: type(x) == datetime.datetime, load_sheet['Date']))[-1]
    
    my_sheet = pd.read_excel(io='data/Contrarian_Indicators.xlsx', sheet_name='AAII')
    my_sheet_lday = my_sheet['Date'].iloc[-1]    
    
    while load_sheet_lday > my_sheet_lday:
        my_sheet_lday += pd.Timedelta(days=7)
        add_data = load_sheet[load_sheet['Date'] == my_sheet_lday].values
        book = xw.Book(fullname='data/Contrarian_Indicators.xlsx')
        sheet = book.sheets['AAII']
        row = int(sheet.range('A1').end('down').row)+1
        sheet.range('A'+str(row)).value = add_data[0]
        sheet.range('E'+str(row)).formula = f'=B{row}-D{row}'
        sheet.range('F'+str(row)).formula = f'=AVERAGE(B{row-7}:B{row})'
        sheet.range('G'+str(row)).formula = f'=AVERAGE(D{row-7}:D{row})'
        book.save()
    
    if load_sheet_delete == True:
        os.remove(r'C:\Users\runru\Downloads\sentiment.xls')
        
        
def get_naaim():
    book = xw.Book('data/Contrarian_Indicators.xlsx')
    sheet = book.sheets['NAAIM']
    last_row = int(sheet.range('A1').end('down').row)
    sheet_lday = sheet.range(f'A{last_row}').value.date()
    
    text = web_scrape(url='https://www.naaim.org/programs/naaim-exposure-index/',
                      css_selector='#surveydata > tbody')
    data = pd.DataFrame(data=[row.split(' ') for row in text.split('\n')][1:],
                        columns=sheet.range('A1:H1').value)
    data['Date'] = [day.date() for day in pd.to_datetime(data['Date'])]
    
    add_data = data[data['Date'] > sheet_lday].sort_values(by=['Date'])
    sheet.range(f'A{last_row+1}').value = add_data.values
    book.save()
    print('Added Data:\n', add_data)        
        
def get_sp500_per():
    """
    Reference:
    S&P500 PER -> https://www.multpl.com/s-p-500-pe-ratio/table/by-month
    S&P500 Shiller -> https://www.multpl.com/shiller-pe/table/by-month
    """
    # Normal PER
    N_PER = web_scrape(url='https://www.multpl.com/s-p-500-pe-ratio/table/by-month',
                       css_selector='#datatable > tbody')
    data_past1 = list(i.replace(' estimate', '').replace(',', '') for i in N_PER.split('\n')[1:14])
    data_past1 = [item.split(' ') for item in data_past1]

    # Shiller PER
    S_PER = web_scrape(url='https://www.multpl.com/shiller-pe/table/by-month',
                       css_selector='#datatable > tbody')
    data_shiller = list(i.replace(' estimate', '').replace(',', '') for i in S_PER.split('\n')[1:14])
    
    # change str to int or float
    for item, shiller in zip(data_past1, data_shiller):
        item[0] = time.strptime(item[0], '%b').tm_mon # month
        item[1:3] = [int(x) for x in item[1:3]] # day and year
        item[-1] = float(item[-1]) # PER
        item.append(float(shiller.split(' ')[-1])) # add shiller PER 
    
    # create DataFrame
    df = pd.DataFrame(data_past1, columns=['M', 'D', 'Y', 'N_PER', 'S_PER'])    
    # add new column 'Date'
    df.insert(loc=0, column='Date', 
              value=[datetime.date(y, m, d) for (y,m,d) in zip(df.Y, df.M, df.D)])
    # sorted by 'Date'
    df.sort_values('Date', inplace=True, ignore_index=True)
    # drop unnessary columns
    df.drop(['M', 'D', 'Y'], axis=1, inplace=True)
    
    # read excel file as pd.DataFrame
    original = pd.read_excel(io='data/Contrarian_Indicators.xlsx',
                             sheet_name='SP500_PER') 
    # get location
    location = original[original['Date'] == df['Date'][0].strftime("%Y-%m-%d")].index
    
    # open excel book
    book = xw.Book(fullname='data/Contrarian_Indicators.xlsx')
    sheet = book.sheets['SP500_PER']
    # change/add the value
    sheet.range(f'A{int(location.values)+2}').value = df.values
    book.save()
    
    print('Past 1 year data:')
    return df
        
def get_margin_debt():
    """
    Reference -> https://www.finra.org/investors/learn-to-invest/advanced-investing/margin-statistics
    """
    book = xw.Book('data/Contrarian_Indicators.xlsx')
    sheet = book.sheets['Margin_Debt']
    last_row = int(sheet.range('A1').end('down').row)
    last_date = sheet.range(f'A{last_row}').value.date()
    
    text = web_scrape(url='https://www.finra.org/investors/learn-to-invest/advanced-investing/margin-statistics',
                      css_selector='#block-finra-bootstrap-sass-system-main > div > article > div > div > div:nth-child(4) > div > div > div > div > div > table:nth-child(5) > tbody')
    data = [row.split(' ') for row in text.split('\n')]
    latest_data = data[-1]
    print("Latest data:\n", latest_data)
    
    if latest_data[0] == last_date.strftime('%b-%y'):
        print('\nYour excel sheet is the latest version, or the margin debt data is not updated.\n')
        return
    
    new_date = last_date + relativedelta(months=1)
    sheet.range(f'A{last_row+1}').value = new_date.strftime('%b %Y')
    sheet.range(f'B{last_row+1}').value = latest_data[1]
    book.save()
    
    
def web_scrape(url: str, css_selector: str):
    """
    Extract data and return it as a text format
    """
    # Automatically download and install the latest chromedriver
    #chromedriver_autoinstaller.install()
    
    path_to_chromedriver = 'C:/Users/runru/Analysis_Data_Python/Finance/chromedriver.exe'
    service = webdriver.ChromeService(path_to_chromedriver)
    driver = webdriver.Chrome(options, service)
        
    driver.implicitly_wait(10)
    driver.get(url)
    element = driver.find_element(by=By.CSS_SELECTOR, value=css_selector)
    text = element.text
    time.sleep(5)
    driver.close()
    return text