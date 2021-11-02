# -*- coding: utf-8 -*-
"""
Performance class which manages performance writing in Excel files

@author: Anco
"""
from workbook import Workbook

class Performance:
    
    def __init__(self, type_perf, data_open, data_close, column, sheet):
        self.type_perf = type_perf
        self.date_open = data_open
        self.data_close = data_close
        self.column = column 
        self.sheet = sheet
        
        # Write performance in workbook 
        # Make a difference between negative and positive performance 
        perf_list = list()
        row = 1 
        if(type_perf == 'price'): 
            for i in range(len(data_open)):
                perf_list.append(data_close[i] - data_open[i])
                if(perf_list[i] > 0):
                    sheet.write(row, column, perf_list[i], Workbook.green_currency_format)
                    row += 1 
                else: 
                    sheet.write(row, column, perf_list[i], Workbook.red_currency_format)
                    row += 1 
        else:
            for i in range(len(data_open)):
                perf_list.append((((data_close[i]*100)/data_open[i])-100)/100)
                if(perf_list[i] > 0):
                    sheet.write(row, column, perf_list[i], Workbook.pourc_green_format)
                    row += 1 
                else: 
                    sheet.write(row, column, perf_list[i], Workbook.pourc_red_format)
                    row += 1