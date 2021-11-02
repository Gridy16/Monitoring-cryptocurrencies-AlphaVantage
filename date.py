# -*- coding: utf-8 -*-
"""
Date class for the column Date of the excel file

@author: Anco
"""
import datetime
from workbook import Workbook

class Date: 
    
    def __init__(self, time, column, sheet): 
        
        self.time = time 
        self.column = column 
        self.sheet = sheet 
        
        # Depending on time scale, create date list and write it in worksheet
        base = datetime.datetime.today()
        numdays = 1000
        row = 1
        if(time.__eq__('daily')): 
            date_list = [base - datetime.timedelta(days=x) for x in range(numdays)]
            for date_list in (date_list):
                sheet.write(row, column, date_list, Workbook.date_format)
                row += 1
        elif(time.__eq__('weekly')):
            date_list = [base - datetime.timedelta(weeks=x) for x in range(numdays)]
            for i in range(len(date_list)):
                sheet.write(row, column, date_list[i], Workbook.date_format)
                row += 1
        else: 
            month_list = list()
            month_list.append(datetime.datetime.now().strftime("%B %Y"))
            for i in range(33):
                month_list.append((datetime.datetime.now() - (i+1)*datetime.timedelta(weeks=4, days=3)).strftime("%B %Y"))
                sheet.write(row, column, month_list[i], Workbook.right_format)
                row += 1
                
                
                
                