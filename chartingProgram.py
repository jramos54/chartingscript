# -*- coding: utf-8 -*-
"""
Created on Mon May  9 20:56:22 2022

@author: zekem
"""

from stockstats import StockDataFrame as Sdf
import yfinance as yf
import plotly.graph_objs as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import requests
import numpy as np
import pandas as pd
import win32com.client as win32
import warnings
#from pandas.core.common import SettingWithCopyWarning
from fpdf import FPDF

#warnings.simplefilter(action="ignore", category=SettingWithCopyWarning)
# proxies = {"http": "http://MXS0MTF:hksyt0mat!!@webproxygo.fpl.com:8080","https": "https://MXS0MTF:hksyt0mat!!@webproxygo.fpl.com:8080",}
# defining style color
colors = {"background": "#ffFFFF", "text": "#000000"}

def get_symbol(symbol):
    x = yf.Ticker(symbol).info
    return x['longName']


def graph_generator(ticker='IWF', ticker_2='IWD', ticker_name='X', option='Relative', period='1wk', sma_1=10, sma_2=6):

# loading data
    start_date = datetime.now().date() - timedelta(days=20 * 365)
    end_data = datetime.now().date()



    # logic for number of periods based on the period
    if period == '1wk':
        time_intervals = 52
    elif period == '1d':
        time_intervals = 252
    else: 
        time_intervals = 12

    # assign the combined data to a dataframe called 'df' depending on whether user has chosen 'relative' or absolute
    if option == "Relative":
        # set future title of chart
        chart_title = yf.Ticker(ticker).info['longName'] + " vs. " + yf.Ticker(ticker_2).info['longName'] + " Performance"
 
        # get the data 
        df_1 = yf.Ticker(ticker).history(period='10y', interval='1wk')
        df_2 = yf.Ticker(ticker_2).history(period='10y', interval='1wk')
        df_1.dropna(inplace=True)
        df_2.dropna(inplace=True)
        df = df_1
        # print(df_1.to_csv('dataframe_1.csv'))
        # print(df_2.to_csv('dataframe_2.csv'))
        # create the Close column
        print("Duplicated Indexes:")
        print(df[df.index.duplicated()])
        df.rename(columns={"Close": "old_Close"})
        df['Close'] = df_1['Close']/df_2['Close']

        # create the percent change column
        df['percent_change'] = df['Close'].pct_change(periods=52)+1
        
    if option == "Absolute":
        
        chart_title = yf.Ticker(ticker).info['longName'] + " Performance"

        # get the data 
        df = yf.Ticker(ticker).history(period='10y', interval='1wk')
        
        # create the percent change column
        df['percent_change'] = df['Close'].pct_change(periods=52)+1
        
    # lowercase all column names
    df.columns = df.columns.str.lower()

    # create stock stats
    stock = Sdf(df)

    std_dev = 0

    # logic for std dev number of periods
    if period == '1wk': std_dev = 104
    elif period == '1d': std_dev = 504
    else: std_dev = 2

    # rolling sma with std deviation bands
    sma_1_series = df.percent_change.rolling(int(sma_1)).mean()
    sma_2_series = sma_1_series.rolling(int(sma_2)).mean()
    std_dev_below = df.percent_change.rolling(std_dev).mean() - df.percent_change.rolling(std_dev).std()
    std_dev_above = df.percent_change.rolling(std_dev).mean() + df.percent_change.rolling(std_dev).std()

    df['sma_1_series'] = sma_1_series
    df['sma_2_series'] = sma_2_series
    df['std_dev_below'] = std_dev_below
    df['std_dev_above'] = std_dev_above

    #logic for signals
    df['signal'] = np.where((sma_1_series > sma_2_series) & (sma_1_series < std_dev_below), 'long',
        (np.where((sma_1_series < sma_2_series) & (sma_1_series > std_dev_above), 'short', 'hold')))

    # logic for buys column
    df['transactions'] = np.where((df.signal == 'long') & (df.signal.shift(1) == 'hold'), 'buy', 
        (np.where((df.signal == 'short') & (df.signal.shift(1) == 'hold'), 'sell', None)))

    # create a column to keep for recording only when buys would occur, then create one with the transaction filled down to simulate 'position'
    df['transactions_not_filled'] = df['transactions']
    df['transactions'] = df['transactions'].ffill()

    #  create a column that reads as 'True' where the position flips
    df['transactions_shifted'] = df['transactions'].shift(1)
    df['transactions_compare'] = df['transactions'] != df['transactions_shifted']

    # if the row is titled buy and the transaction flips for that row, add to buys column
    df['buys'] = np.where((df.transactions == 'buy') & (df.transactions_compare == True), sma_1_series, '')
    ''' ALSO FILTER SELLS AND SEE IF THE LAST BUY IS BEFORE THE LAST SELL'''
    # df['buys'] = np.where((df.transactions == 'buy') & (df.std_dev_below.shift(1) >= df.sma_1_series.shift(1)) & (df.std_dev_below < df.sma_1_series), sma_1_series, '')

    ''' ALSO FILTER BUYS AND SEE IF THE LAST SELL IS BEFORE THE LAST BUY'''
    # if the row is titled sell and the transaction flips for that date, add to sells column
    df['sells'] = np.where((df.transactions == 'sell') & (df.transactions_compare == True), sma_1_series, '')
    # df['sells'] = np.where((df.transactions == 'sell') &(df.std_dev_above.shift(1) <= df.sma_1_series.shift(1)) & (df.std_dev_above > df.sma_1_series), sma_1_series, '')

    # create a table of when the buys and sells have occured (for doing calculations)
    trades = df.loc[(df['sells'] != '') | (df['buys'] != '')]
    trades['duplicates'] = trades.transactions != trades.transactions.shift(1)
    trades = trades.loc[trades.duplicates == True]
    trades['percentage_pickup'] = trades['close'].pct_change()

    # # export the transactions
    # trades.to_csv(ticker + ' trade log.csv')
    # df.to_csv(ticker + ' all data.csv')

    fig_3 = make_subplots(rows=3, cols=1,
                shared_xaxes=True,
                vertical_spacing=0.02)

    fig_3 = go.Figure(
        data=[
            go.Scatter(
                x=list(sma_1_series.index), y=list(sma_1_series), name=str(sma_1) + " day SMA",
                line=dict(color="blue")
            ),
            go.Scatter(
                x=list(sma_2_series.index), y=list(sma_2_series), name=str(sma_2) + " day SMA",
                line=dict(color="orange")
            ),
            go.Scatter(
                x=list(std_dev_above.index), y=list(std_dev_above), name= "+1 Std Dev.",
                line=dict(color="lightsalmon", width=2, dash='dot')
            ),
            go.Scatter(
                x=list(std_dev_below.index), y=list(std_dev_below), name= "-1 Std Dev.",
                line=dict(color="lightgreen", width=2, dash='dot')
            ),
            go.Scatter(
                x=list(trades.index), y=list(trades.buys), name= "Buys", mode='markers',marker=dict(size=15, color='green')
            ),
            go.Scatter(
                x=list(trades.index), y=list(trades.sells), name= "Sells", mode='markers', marker=dict(size=15, color='red')
            ),
        ],
        layout={
            "title": chart_title,
            "height": 1000,
            "showlegend": True,
            "plot_bgcolor": colors["background"],
            "paper_bgcolor": colors["background"],
            "font": {"color": colors["text"]},
        },
    )
    
    fig_3.update_xaxes(
        rangeslider_visible=False,
        rangeselector=dict(
            activecolor="blue",
            bgcolor=colors["background"],
            buttons=list(
                [
                    dict(count=7, label="10D",
                            step="day", stepmode="backward"),
                    dict(
                        count=15, label="15D", step="day", stepmode="backward"
                    ),
                    dict(
                        count=1, label="1m", step="month", stepmode="backward"
                    ),
                    dict(
                        count=3, label="3m", step="month", stepmode="backward"
                    ),
                    dict(
                        count=6, label="6m", step="month", stepmode="backward"
                    ),
                    dict(count=1, label="1y", step="year",
                            stepmode="backward"),
                    dict(count=5, label="5y", step="year",
                            stepmode="backward"),
                    dict(count=1, label="YTD",
                            step="year", stepmode="todate"),
                    dict(step="all"),
                ]
            ),
        ),
    )

    
    fig_3.write_image(datetime.today().strftime("%m-%d-%Y") + ' ' + ticker + ' ' + ticker_2 + r" buys and sells.png", width=1200, height=900)
    image_link = r'C:\Users\joelr\Documents\codingFiles\chartingProgram\pics' + "\\" +  datetime.today().strftime("%m-%d-%Y") + ' ' + ticker + ' ' + ticker_2 + r" buys and sells.png"
    # image_link = r'\\Jbxsf70\jbfin$\TRUST\PUBLIC\1 TFI Sharepoint\2021\Projects\Equity Tactical Allocation\Python Scripts\ETF Email' + "\\" +  datetime.today().strftime("%m-%d-%Y") + ' ' + ticker + ' ' + ticker_2 + r" buys and sells.png"
    print(image_link)
    return image_link

# Other
item_1 = graph_generator('IWF','IWD', get_symbol('IWD'), 'Relative', '1wk', 10, 6)
# item_2 = graph_generator('IWO','IWN', get_symbol('IWN'), 'Relative', '1wk', 10, 6)
# item_3 = graph_generator('IWF','IWO', get_symbol('IWO'), 'Relative', '1wk', 10, 6)
# item_4 = graph_generator('IWD','IWN', get_symbol('IWN'), 'Relative', '1wk', 10, 6)
# item_5 = graph_generator('IWB','IWM', get_symbol('IWM'), 'Relative', '1wk', 10, 6)
# item_6 = graph_generator('SPY','EFA', get_symbol('EFA'), 'Relative', '1wk', 10, 6)
# item_7 = graph_generator('SPY','EEM', get_symbol('EEM'), 'Relative', '1wk', 10, 6)
# item_8 = graph_generator('QQQ','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_9 = graph_generator('PSCT','QQQ', get_symbol('QQQ'), 'Relative', '1wk', 10, 6)
# item_10 = graph_generator('VGK','ACWI', get_symbol('ACWI'), 'Relative', '1wk', 10, 6)
# item_11 = graph_generator('IEUS','ACWI', get_symbol('ACWI'), 'Relative', '1wk', 10, 6)
# item_12 = graph_generator('EWJ','ACWI', get_symbol('ACWI'), 'Relative', '1wk', 10, 6)
# item_13 = graph_generator('AAXJ','ACWI', get_symbol('ACWI'), 'Relative', '1wk', 10, 6)
# item_14 = graph_generator('ILF','ACWI', get_symbol('ACWI'), 'Relative', '1wk', 10, 6)
# item_40 = graph_generator('ICF','SPY', get_symbol('ACWI'), 'Relative', '1wk', 10, 6)

# # BONDS
# item_33 = graph_generator('VCIT','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_34 = graph_generator('HYG','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_35 = graph_generator('CWB','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_36 = graph_generator('AGG','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_37 = graph_generator('SRLN','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)


# # US SECTORS
# item_15 = graph_generator('XLK','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_16 = graph_generator('XLE','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_17 = graph_generator('XLI','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_18 = graph_generator('XLY','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_19 = graph_generator('XLB', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_20 = graph_generator('IYZ', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_21 = graph_generator('XLF', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_22 = graph_generator('XLV', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_23 = graph_generator('XLP', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_24 = graph_generator('XLU', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)

# # THEMES

# item_25 = graph_generator('NXTG', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_26 = graph_generator('URA', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_27 = graph_generator('USO', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_28 = graph_generator('ICLN', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_29 = graph_generator('LIT', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_30 = graph_generator('NEE', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_31 = graph_generator('IBB', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_32 = graph_generator('IHE', 'SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_38 = graph_generator('IEZ','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)
# item_39 = graph_generator('XOP','SPY', get_symbol('SPY'), 'Relative', '1wk', 10, 6)



# input_file = "\\Jbxsf70\jbfin$\TRUST\PUBLIC\1 TFI Sharepoint\2021\Projects\Equity Tactical Allocation\Weekly PDFs\example.pdf"
# output_file = r"\\Jbxsf70\jbfin$\TRUST\PUBLIC\1 TFI Sharepoint\2021\Projects\Equity Tactical Allocation\Weekly PDFs\ETF email {}.pdf".format(datetime.today().strftime("%m-%d-%Y"))
# picture_list = 'item_1,item_2,item_3,item_4,item_5,item_6,item_7,item_8,item_9,item_10,item_11,item_12,item_13,item_14,item_15,item_16,item_17,item_18,item_19,item_20,item_21,item_22,item_23,item_24,item_25,item_26,item_27,item_28,item_29,item_30,item_31,item_32'
# pdf = FPDF()

# for files in picture_list:
#     pdf.add_page()
#     pdf.image(files)

# pdf.output(output_file, 'F')

"""
# create an instance of outlook
outlook = win32.Dispatch('outlook.application')

# send an email to NEER business leaders
mail = outlook.CreateItem(0)
mail.To = "zekemaki@comcast.net" 
mail.Subject = "The Weekly Factor ETF Download | " + datetime.today().strftime("%m-%d-%Y")


attachment_2 = mail.Attachments.Add(item_1)
attachment_2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_1")
# attachment_3 = mail.Attachments.Add(item_2)
# attachment_3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_2")
# attachment_4 = mail.Attachments.Add(item_3)
# attachment_4.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_3")
# attachment_5 = mail.Attachments.Add(item_4)
# attachment_5.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_4")
# attachment6 = mail.Attachments.Add(item_5)
# attachment6.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_5")
# attachment7 = mail.Attachments.Add(item_6)
# attachment7.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_6")
# attachment8 = mail.Attachments.Add(item_7)
# attachment8.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_7")
# attachment9 = mail.Attachments.Add(item_8)
# attachment9.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_8")
# attachment10 = mail.Attachments.Add(item_9)
# attachment10.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_9")
# attachment11 = mail.Attachments.Add(item_10)
# attachment11.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_10")
# attachment12 = mail.Attachments.Add(item_11)
# attachment12.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_11")
# attachment13 = mail.Attachments.Add(item_12)
# attachment13.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_12")
# attachment14 = mail.Attachments.Add(item_13)
# attachment14.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_13")
# attachment15 = mail.Attachments.Add(item_14)
# attachment15.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_14")
# attachment16 = mail.Attachments.Add(item_15)
# attachment16.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_15")
# attachment17 = mail.Attachments.Add(item_16)
# attachment17.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_16")
# attachment18 = mail.Attachments.Add(item_17)
# attachment18.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_17")
# attachment19 = mail.Attachments.Add(item_18)
# attachment19.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_18")
# attachment20 = mail.Attachments.Add(item_19)
# attachment20.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_19")
# attachment22 = mail.Attachments.Add(item_20)
# attachment22.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_20")
# attachment23 = mail.Attachments.Add(item_21)
# attachment23.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_21")
# attachment24 = mail.Attachments.Add(item_22)
# attachment24.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_22")
# attachment25 = mail.Attachments.Add(item_23)
# attachment25.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_23")
# attachment26 = mail.Attachments.Add(item_24)
# attachment26.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_24")
# attachment27 = mail.Attachments.Add(item_25)
# attachment27.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_25")
# attachment28 = mail.Attachments.Add(item_26)
# attachment28.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_26")
# attachment29 = mail.Attachments.Add(item_27)
# attachment29.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_27")
# attachment30 = mail.Attachments.Add(item_28)
# attachment30.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_28")
# attachment31 = mail.Attachments.Add(item_29)
# attachment31.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_29")
# attachment32 = mail.Attachments.Add(item_30)
# attachment32.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_30")
# attachment33 = mail.Attachments.Add(item_31)
# attachment33.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_31")
# attachment34 = mail.Attachments.Add(item_32)
# attachment34.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_32")

# attachment35 = mail.Attachments.Add(item_33)
# attachment35.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_33")
# attachment36 = mail.Attachments.Add(item_34)
# attachment36.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_34")
# attachment37 = mail.Attachments.Add(item_35)
# attachment37.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_35")
# attachment38 = mail.Attachments.Add(item_36)
# attachment38.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_36")
# attachment39 = mail.Attachments.Add(item_37)
# attachment39.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_37")
# attachment40 = mail.Attachments.Add(item_38)
# attachment40.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_38")
# attachment41 = mail.Attachments.Add(item_39)
# attachment41.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_39")
# attachment42 = mail.Attachments.Add(item_40)
# attachment42.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "item_40")





mail.HTMLBody = "<html><body><h1>Main Charts</h1><img src=""cid:item_1""><img src=""cid:item_2""><img src=""cid:item_3""><img src=""cid:item_4""><img src=""cid:item_5""><img src=""cid:item_6""><img src=""cid:item_7""><img src=""cid:item_8""><img src=""cid:item_9""><img src=""cid:item_10""><img src=""cid:item_11""><img src=""cid:item_12""><img src=""cid:item_13""><img src=""cid:item_14""><img src=""cid:item_40""><h1>Bonds</h1><img src=""cid:item_33""><img src=""cid:item_34""><img src=""cid:item_35""><img src=""cid:item_36""><img src=""cid:item_37""><h1>US Sectors</h1><img src=""cid:item_15""><img src=""cid:item_16""><img src=""cid:item_17""><img src=""cid:item_18""><img src=""cid:item_19""><img src=""cid:item_20""><img src=""cid:item_21""><img src=""cid:item_22""><img src=""cid:item_23""><img src=""cid:item_24""><h1>Themes</h1><img src=""cid:item_25""><img src=""cid:item_26""><img src=""cid:item_27""><img src=""cid:item_28""><img src=""cid:item_29""><img src=""cid:item_30""><img src=""cid:item_31""><img src=""cid:item_32""><img src=""cid:item_38""><img src=""cid:item_39""></body></html>"
mail.send

"""