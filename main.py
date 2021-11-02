# -*- coding: utf-8 -*-
"""
Monitoring cryptocurrencies performance in Excel through AlphaVantage API 

@author : Anco
"""
from excelFile import ExcelFile
from workbook import Workbook

"""
You choose here what cryptocurrencies you want to have in your Excel 

Here's Bitcoin and Ethereum performance through different timescales
"""

worksheet = Workbook.workbook.add_worksheet("BTC Perf Daily")
ExcelFile('daily', 'BTC', worksheet)

worksheet2 = Workbook.workbook.add_worksheet("BTC Perf Weekly")
ExcelFile('weekly', 'BTC', worksheet2)

worksheet3 = Workbook.workbook.add_worksheet("BTC Perf Monthly")
ExcelFile('monthly', 'ETH', worksheet3)

worksheet4 = Workbook.workbook.add_worksheet("ETH Perf Daily")
ExcelFile('daily', 'ETH', worksheet4)

worksheet5 = Workbook.workbook.add_worksheet("ETH Perf Weekly")
ExcelFile('weekly', 'ETH', worksheet5)

worksheet6 = Workbook.workbook.add_worksheet("ETH Perf Monthly")
ExcelFile('monthly', 'ETH', worksheet6)

Workbook.workbook.close()
