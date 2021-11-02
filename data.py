# -*- coding: utf-8 -*-
"""
Data class which manages data writing in Excel files

@author: Anco
"""
from workbook import Workbook

class Data: 
    
    def __init__(self, data, column, sheet): 
        self.data = data
        self.column = column 
        self.sheet = sheet
        
        # Write data in worksheet
        row = 1 
        for data in (data):
            sheet.write(row, column, data, Workbook.currency_format)
            row += 1