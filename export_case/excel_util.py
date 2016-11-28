# -*- coding=utf-8 -*-
# author: yanyang.xie@gmail.com

import os
import string

import xlrd
from xlutils import copy
import xlwt
from xlwt.ExcelFormula import Formula


class ExcelUtil:
    '''
    Excel operation utility.
    @param excel_file_path: excel file absolute path. If none, then should be create an excel, otherwise read excel from local.
    @param formatting_info: whether to keep format for original excel file
    @param on_demand: just use it for larger excel file. just read sheet while use 
    '''

    def __init__(self, excel_file_path=None, formatting_info=True, on_demand=False):
        if excel_file_path is None:
            self.xlrd_workbook = None
            self.xlwt_workbook = xlwt.Workbook()
        else:
            self.excel_file_path = excel_file_path
            self.xlrd_workbook = xlrd.open_workbook(excel_file_path, formatting_info, on_demand)
            self.xlwt_workbook = copy.copy(self.xlrd_workbook)
        
        self.work_sheets = {}
    
    #----------------read, using xlrd------------------------------
    # get sheet number
    def get_sheet_count(self):
        if self.xlrd_workbook:
            return len(self.xlrd_workbook.sheets())
    
    # get work sheet object(sheet.Sheet)
    def get_sheet_by_name(self, sheet_name):
        return self.xlrd_workbook.sheet_by_name(sheet_name)

    # get sheet by index
    def get_sheet_name_by_index(self, index):
        if self.xlrd_workbook:
            return self.xlrd_workbook.sheet_names()[index]

    # get index by sheet name
    def get_index_by_sheet_name(self, sheet):
        if self.xlrd_workbook:
            return self.xlrd_workbook.sheet_names().index(sheet)

    # get max rows in one sheet
    def get_sheet_max_rows(self, sheet_name):
        value = 0
        if self.xlrd_workbook:
            work_sheet = self.get_sheet_by_name(sheet_name)
            if work_sheet != None:
                value = work_sheet.nrows
        return value
    
    # get max columns in one sheet
    def get_sheet_max_columns(self, sheet_name):
        value = 0
        if self.xlrd_workbook:
            work_sheet = self.xlrd_workbook.sheet_by_name(sheet_name)
            if work_sheet != None:
                value = work_sheet.ncols
        return value

    # get cell content by sheet name and row + col
    def get_cell_value(self, sheet_name, row, col):
        value = None
        if self.xlrd_workbook:
            work_sheet = self.get_sheet_by_name(sheet_name)
            if work_sheet != None:
                value = work_sheet.cell(row, col).value
                if value is not None:
                    try:
                        v = int(value)
                        if v == value:
                            value = str(v)
                    except:
                        value = string.strip(value)
        return value

    def get_row_values(self, sheet_name, row, start_colx=0, end_colx=None):
        if self.xlrd_workbook:
            work_sheet = self.get_sheet_by_name(sheet_name)
            return work_sheet.row_values(row, start_colx, end_colx)
    
    def get_column_values(self, sheet_name, col, start_rowx=0, end_rowx=None):
        if self.xlrd_workbook:
            work_sheet = self.get_sheet_by_name(sheet_name)
            return work_sheet.col_values(col, start_rowx, end_rowx)

    # get cell values by range in one row. Range should be as 0:0:8
    def get_row_values_in_range(self, sheet_name, c_range="0:0:8"):
        values = []
        row, start, end = c_range.split(':')
        if self.xlrd_workbook:
            work_sheet = self.get_sheet_by_name(sheet_name)
            if work_sheet != None:
                values = work_sheet.row_values(int(row), int(start), int(end))
        return values
    
    # get cell values by range in one column. Range should be as 0:0:8
    def get_column_values_in_range(self, sheet_name, c_range="0:0:8"):
        value = None
        row, start, end = c_range.split(':')
        if self.xlrd_workbook:
            work_sheet = self.get_sheet_by_name(sheet_name)
            if work_sheet != None:
                values = work_sheet.col_values(int(row), int(start), int(end))
                value = string.join([v for v in values if v.strip() != ''])
        return value
    
    #----------------write, using xlwt------------------------------
    # just for modify a excel, using sheet name to get index and then using the index to get work sheet in xlwt_workbook
    def get_xlwt_work_sheet(self, sheet_name):
        if self.xlWBook:
            index = self.get_index_by_sheet_name(sheet_name)
            work_sheet = self.xlWBook.get_sheet(index)
            return work_sheet
    
    # add a new sheet
    def add_sheet(self, sheet_name, cell_overwrite_ok=True):
        if self.xlwt_workbook:
            ws = self.xlwt_workbook.add_sheet(sheet_name, cell_overwrite_ok)
            self.work_sheets[sheet_name] = ws
            return ws
    
    # set cell value and pattern
    # style = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;')
    # if pattern is None, xlwt will set default style to cell, not use the row style or column style
    def set_cell_value(self, work_sheet, row, col, value, pattern=None):
        if pattern is not None:
            work_sheet.write(row, col, value, xlwt.easyxf(pattern))
        else:
            work_sheet.write(row, col, value)
    
    # set cell value in special range
    def set_merged_cell_value(self, work_sheet, row_start, row_end, column_start, column_end, value, pattern=None):
        work_sheet.write_merge(row_start, row_end, column_start, column_end, value, xlwt.easyxf(pattern))
    
    def set_link_cell_value(self, work_sheet, row, col, link, link_text, pattern=None):
        value = Formula('HYPERLINK("%s";"%s")' % (link, link_text))
        work_sheet.write(row, col, value, xlwt.easyxf(pattern))
    
    def set_formula_cell_value(self, work_sheet, row, col, value, pattern=None):
        value = Formula('%s' % (value))
        work_sheet.write(row, col, value, xlwt.easyxf(pattern))
    
    # x, y is the relative position from the upper left corner in row+column
    # scale_x and scale_y is to zoom in or out original image
    # xlwt only support 24bit bmp file, if you need jpg or png, please use xlsxwriter
    def insert_24bit_bmp_image(self, work_sheet, image_file_path, row, column, x=0, y=0, scale_x=1, scale_y=1):
        if os.path.exists(image_file_path):
            work_sheet.insert_bitmap(image_file_path, row, column, x, y, scale_x, scale_y)
    
    # set column style
    def set_column_style(self, work_sheet, column_index, column_width=None, pattern=None):
        if column_width is not None:
            work_sheet.col(column_index).width = column_width
            
        if pattern is not None:
            work_sheet.col(column_index).set_style(xlwt.easyxf(pattern))
    
    # set row style
    def set_row_style(self, work_sheet, row_index, pattern=None):
        if pattern is not None:
            work_sheet.row(row_index).set_style(xlwt.easyxf(pattern))
            
    # save excel contents to local file
    def save(self, new_file_path=None):
        if new_file_path is not None:
            self.excel_file_path = new_file_path

        if self.xlwt_workbook and self.excel_file_path is not None:
            self.xlwt_workbook.save(self.excel_file_path)

def _testExport():
    t_file = 'e:/1.xls'
    if os.path.exists(t_file):
        os.remove(t_file)
    
    excel_util = ExcelUtil()
    work_sheet = excel_util.add_sheet(u'Basic test')
    excel_util.set_column_style(work_sheet, column_index=1, column_width=5000)
    excel_util.set_column_style(work_sheet, column_index=2, column_width=10000)
    excel_util.set_column_style(work_sheet, 3, pattern='font: color-index blue')
    excel_util.set_row_style(work_sheet, 3, pattern='font: color-index yellow')
    excel_util.set_cell_value(work_sheet, 3, 3, u'3,3')
    excel_util.set_cell_value(work_sheet, 1, 1, u'1,1', pattern='font: color-index blue')
    excel_util.set_cell_value(work_sheet, 2, 2, u'2,2wo我们', pattern='font: color-index red, bold on; align: vertical center, horizontal center;')
    excel_util.set_merged_cell_value(work_sheet, 5, 8, 1, 2, u'5-8, 1-2', pattern='font: color-index blue')
    # excel_util.insert_24bit_bmp_image(work_sheet, 'e:/1.bmp', 9, 9, 5, 10)
    excel_util.set_link_cell_value(work_sheet, 0, 0, 'http://www.python.org', 'python', pattern='font: color-index blue')

    work_sheet2 = excel_util.add_sheet(u'Formula test')
    excel_util.set_cell_value(work_sheet2, 0, 0, 10)
    excel_util.set_cell_value(work_sheet2, 0, 1, 20)
    excel_util.set_formula_cell_value(work_sheet2, 1, 0, 'A1/B1')
    excel_util.set_formula_cell_value(work_sheet2, 1, 1, 'sum(1,2,3)')
    excel_util.save(t_file)
    print 'done'


if __name__ == '__main__':
    _testExport()

'''
easyxf(function) 创建 XFStyle instance，格式控制

expression syntax: <element>:(<attribute> <value>, <attribute> <value>, ); <element>:(<attribute> <value>,);

font      - bold          - True or False
          - colour        - {colour}
          - italic        - True or False
          - name          - name of the font, Arial
          - underline     - True or False
alignment - direction     - general , lr, rl
          - horizontal    - general , left, center, right, filled 
          - vertical      - bottom , top, center, justified, distributed
          - shrink_to_fit - True or False
borders   - left          - an integer width between 0 and 13
          - right         - an integer width between 0 and 13
          - top           - an integer width between 0 and 13
          - bottom        - an integer width between 0 and 13
          - diag          - an integer width between 0 and 13
          - left_colour   - {colour}*, automatic colour
          - right_colour  - {colour}*,  automatic colour
          - ...
pattern   - back_color    - {colour}*,  automatic colour
          - fore_colour   - {colour}*,  automatic colour
          - pattern       - none , solid, fine_dots, sparse_dots
fore_colour = ['aqua','black','blue','blue_gray','bright_green','brown','coral','cyan_ega','dark_blue','dark_blue_ega','dark_green','dark_green_ega','dark_purple','dark_red',
                'dark_red_ega','dark_teal','dark_yellow','gold','gray_ega','gray25','gray40','gray50','gray80','green','ice_blue','indigo','ivory','lavender',
                'light_blue','light_green','light_orange','light_turquoise','light_yellow','lime','magenta_ega','ocean_blue','olive_ega','olive_green','orange','pale_blue','periwinkle','pink',
                'plum','purple_ega','red','rose','sea_green','silver_ega','sky_blue','tan','teal','teal_ega','turquoise','violet','white','yellow']
'''

'''
#定义写入单元格字体格式，包含中文时必须带u
title=easyxf(u'font:name 仿宋,height 240 ,colour_index black, bold on, italic off; align: wrap on, vert centre, horiz center;pattern: pattern solid, fore_colour light_orange;') #字体黑色加粗，自动换行、垂直居中、水平居中,背景色橙色  
normal=easyxf('font:colour_index black, bold off, italic off; align: wrap on, vert centre, horiz left;') #字体黑色不加粗，自动换行、垂直居中、水平居左  
warn=easyxf('font:colour_index red, bold off, italic off; align: wrap on, vert centre, horiz left;')#字体红色，自动换行，垂直居中，水平居中 
'''
