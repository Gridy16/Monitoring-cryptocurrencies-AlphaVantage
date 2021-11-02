# -*- coding: utf-8 -*-
"""
Worbook class which defines global features of Excel files 

@author: Anco
"""
import xlsxwriter

class Workbook: 
    
    workbook = xlsxwriter.Workbook('DataCrypto.xlsx')
    
    # Workbook formats 
    date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    
    green_currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    red_currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    green_currency_format.set_font_color('green')
    red_currency_format.set_font_color('red')
    green_currency_format.set_bold()
    red_currency_format.set_bold()
    
    pourc_green_format = workbook.add_format({'num_format': '0.00%'})
    pourc_red_format = workbook.add_format({'num_format': '0.00%'})
    pourc_green_format.set_font_color('green')
    pourc_red_format.set_font_color('red')
    pourc_green_format.set_bold()
    pourc_red_format.set_bold()
    
    center_bg_format = workbook.add_format()
    center_bg_format.set_align('center')
    center_bg_format.set_align('vcenter')
    center_bg_format.set_bg_color('#CFCBCB')
    
    right_format = workbook.add_format()
    right_format.set_align('right')