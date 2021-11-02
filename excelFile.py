# -*- coding: utf-8 -*-
"""
Class ExcelFile which creates the file depending on the cryptocurrency
and the timescale entered by the user 

@author: Anco
"""

# Modules import
from alpha_vantage.cryptocurrencies import CryptoCurrencies
from date import Date
from data import Data
from performance import Performance
from workbook import Workbook 

ALPHAVANTAGE_API_KEY = 'ENTERYOURKEYHERE' # AlphaVantage API key 
cc = CryptoCurrencies(key=ALPHAVANTAGE_API_KEY, output_format='pandas') # Stock in pandas fromat Crypto data from AlphaVantage API 

class ExcelFile:
    
    def __init__(self, timeScale, symbol, worksheet):
        self.timeScale = timeScale
        self.symbol = symbol 
        self.worksheet = worksheet
        
        # Collect data depending on time scale
        if(timeScale.__eq__('daily')):
            _data, meta_data = cc.get_digital_currency_daily(symbol=symbol, market='CNY')
        elif(timeScale.__eq__('weekly')): 
            _data, meta_data = cc.get_digital_currency_weekly(symbol=symbol, market='CNY')
        else: 
            _data, meta_data = cc.get_digital_currency_monthly(symbol=symbol, market='CNY')
        
        # Take only open and close data from data (check AlphaVantage help for more details)
        openData = _data['1b. open (USD)'] 
        closeData = _data['4b. close (USD)']
            
        # Columns of the excel file 
        Date(timeScale, 0, worksheet) 
        Data(openData, 1, worksheet)
        Data(closeData, 2, worksheet)
        Performance('price', openData, closeData, 3, worksheet)
        Performance('pourcentage', openData, closeData, 4, worksheet)
       
        # Columns titles and size
        worksheet.write(0, 0, "Date", Workbook.center_bg_format)
        worksheet.write(0, 1, "" + str(symbol) + " Open (USD)", Workbook.center_bg_format)  
        worksheet.write(0, 2, "" + str(symbol) + " Close (USD)", Workbook.center_bg_format)
        worksheet.write(0, 3, "Performance (USD)", Workbook.center_bg_format)
        worksheet.write(0, 4, "Performance (%)", Workbook.center_bg_format)
        worksheet.set_column('A:F', 22)
        worksheet.set_row(0, 25)
      
        