"""
Class to help build XLSX Spreadsheets using the xlsxwriter library. The class
is pretty basic, and wouldn't really be necessary except for the way that
xlsxwriter handles styling.
Styles in xlsxwriter cannot be created independent of a workbook class, which
means that there isn't a way to create a single library of styles. With this
class however, the styles can be defined in a single place and used everywhere.
There is also functionality to create new styles on the fly if needed.
In addition to the styling, there are methods for creating a sheet and filling
it from a simple list of records and from a complicated object. The fill from
object isn't very generic and should probably be removed. Both of the filling
methods require a column dictionary which describes several aspects of the data
in the column, e.g. header, width, style of the data, etc. Here is an example
of data and a column dictionary:
data = [
    ['hello', 'world', Decimal('.0333'), 1500,
        -4.50, '01/01/2014', {'url':'https://www.test.com', 'tip': 'Test'}],
    ['foo', 'bar', Decimal('.1735'), 30,
        100, '12/15/2000', {'url': 'http://www.google.com', 'string': 'google'}],
    ['spam', 'egg', Decimal('.1'), 18000,
        235543, '06/06/2006', 'http://www.shoutlet.com'],
]
col_dict = {
        0: {'label': 'text 1', 'width': 12, 'style': 'text_style'},
        1: {'label': 'text 2', 'width': 12, 'style': 'text_style'},
        2: {'label': 'Pct 1', 'width': 12, 'style': 'pct_style'},
        3: {'label': 'Int 1', 'width': 8, 'style': 'int_style'},
        4: {'label': 'Currency', 'width': 26, 'style': 'currency_style'},
        5: {'label': 'Date 1', 'width': 12, 'style': 'date_style'},
        6: {'label': 'URL', 'width': 26, 'style': 'url_style'},
        }
--Chris Meyers (cmeyers@zendesk.com) 2017-03-08
"""
import xlsxwriter
import time
from decimal import Decimal
import pandas as pd

class XLSXWorkbook():
    def __init__(self, filename):
        """
        Init for the class. Since the workbook is needed for all other aspects
        of the class, one will be created here.
        filename: the name of the file where the spreadsheet will be written
        """
        self.filename = filename
        self.workbook = xlsxwriter.Workbook(self.filename)
        self.build_default_styles()
    def get_new_worksheet(self, sheetname):
        """
        Add a new sheet to the workbook.
        sheetname: the name of the sheet, should follow spreadsheet naming
            conventions
        """
        sheet = self.workbook.add_worksheet(sheetname)
        return sheet
    def set_style(self, stylename, params):
        """
        Method for adding a style to the workbook. Each style will be a class-
        level attribute.
        stylename: the text name of the style, can be anything
        params: a dictionary of parameters that will be part of the style
        """
        style = self.workbook.add_format(params)
        setattr(self, stylename, style)
    def build_default_styles(self):
        """
        Place to define the standard cell styles. Each style is a dictionary
        of parameters. See the docs for more information:
            https://xlsxwriter.readthedocs.org/working_with_formats.html
        """
        # print('Building default styles')
        text_params = {'align': 'left', 'font_name': 'Helvetica'}
        color_text_params = {'align': 'left','font_name': 'Helvetica', 'font_color':'#5951ff'} # Conditions and Messages & Events
        color_bold_text_params = {'align': 'left','bold': True,'font_name': 'Helvetica',} # conditional operator text - AND or OR
        color_checkboxes_params = {'align': 'center','bold': True,'font_name': 'Helvetica','font_size': 20,'font_color': '#14A139'} # columns contain checkboxes
        hdr_params = {'bold': True,
                'align': 'left',
                'shrink': True,
                'bg_color': '#5951ff', # header background color set to deep purple
                'font_name': 'Helvetica',
                'font_size': 15,
                'font_color': '#ffffff'
                }
        sub_hdr_params = {'bold': True,
                'align': 'center',
                'shrink': True,
                'font_name': 'Helvetica',
                'font_color':'#5951ff',
                'bottom': 5
                }
        bold_params = {'bold': True,'font_name':'Helvetica','align':'left'}
        date_params = {'num_format': 'mm/dd/yyyy'}
        time_params = {'num_format': 'hh:mm'}
        integer_params = {'num_format': '#,##0', 'align': 'right'}
        number_params = {'num_format': '#,##0.00'}
        idnum_params = {'num_format': '###0', 'align': 'right'}
        pct_params = {'num_format': '0.00%'}
        currency_params = {'num_format': '_($#,##0.00_);[Red]_(-$#,##0.00_)'}
        # Expanding the basic parameters with a bold font-weight and a double-
        # lined top border for total rows.
        total_params = {'bold': True,
                'top': 6,
                }
        date_total_params = dict(total_params, **date_params)
        int_total_params = dict(total_params, **integer_params)
        num_total_params = dict(total_params, **number_params)
        pct_total_params = dict(total_params, **pct_params)
        currency_total_params = dict(total_params, **currency_params)
        text_total_params = dict(total_params, **text_params)
        # Need to run all of the param dictionaries through the set_style
        # method. Many ways to do this, could call set_style for each one
        # individually. I just thought this was a bit more compact and less
        # repetitive.
        all_params = {'text_style': text_params,'color_text_style': color_text_params, 
        'color_bold_text_style': color_bold_text_params, 'color_checkboxes':color_checkboxes_params,
                'sub_hdr_style':sub_hdr_params,'hdr_style': hdr_params,
                'bold_style': bold_params, 'total_style': total_params,
                'date_style': date_params, 'time_style': time_params,
                'int_style': integer_params, 'num_style': number_params,
                'pct_style': pct_params, 'currency_style': currency_params,
                'idnum_style': idnum_params,
                'date_tot_style': date_total_params, 'int_tot_style': int_total_params,
                'num_tot_style': num_total_params, 'pct_tot_style': pct_total_params,
                'currency_tot_style': currency_total_params,
                'text_tot_style': text_total_params}
        for stylename, params in all_params.items():
            self.set_style(stylename, params)
    def add_headers(self, sheet, col_dict, multicol_max_length):
        """
        Method for adding header labels to a sheet. Will apply the hdr_style
        formatting to each cell and will set the width. The label and width
        parameters come from the col_dict.
        sheet: a sheet object that has been added to a workbook
        col_dict: a dictionary of meta-data about each column
        """
        for col, metadata in col_dict.items():
            multicol = metadata.get('multicolumn', False)
            if not multicol:
                sheet.set_column(col, col, metadata['width'])
                sheet.write(0, col, metadata['label'], self.hdr_style)
            else:
                for i in range(0, multicol_max_length):
                    new_col = col + i
                    sheet.set_column(new_col, new_col, metadata['width'])
                    sheet.write(0, new_col, metadata['label'], self.hdr_style)
    def add_sub_headers(self, sheet, col_dict, multicol_max_length,row,column):
        """
        Method for adding header labels to a sheet. Will apply the hdr_style
        formatting to each cell and will set the width. The label and width
        parameters come from the col_dict.
        sheet: a sheet object that has been added to a workbook
        col_dict: a dictionary of meta-data about each column
        row: line where to add the header
        column: column to add the header
        """
        for col, metadata in col_dict.items():
            col= col + column
            multicol = metadata.get('multicolumn', False)
            if not multicol:
                sheet.set_column(col, col, metadata['width'])
                sheet.write(row, col, metadata['label'], self.sub_hdr_style)
            else:
                for i in range(0, multicol_max_length):
                    new_col = col + i
                    sheet.set_column(new_col, new_col, metadata['width'])
                    sheet.write(row, new_col, metadata['label'], self.sub_hdr_style)
    def _write_data_to_column(self, sheet, row, col, metadata, data, multicol_max_length):
        """
        Data for a column can have special requirements, e.g. URLs can have
        special formating and a display string. This class-only method deals
        with those requirements.
        sheet: a sheet object that has been added to a workbook.
        style_string: the text of the name of the style, i.e. int_style
        data: the actual data going into the cell, could be a string, number,
            or in the case of a URL, a dictionary
        """
        style_string = metadata['style']
        if 'note' in metadata.keys():
            # print(row,col,metadata)
            sheet.write_comment(0,col,metadata['note'])
        if style_string == 'url_style':
            if isinstance(data, dict):
                # If the url data is a dictionary, that means that it could
                # contain formatting options.
                sheet.write_url(row, col, **data)
            else:
                sheet.write_url(row, col, data)
        elif metadata.get('multicolumn', False):
            for i, val in enumerate(data):
                style = getattr(self, style_string)
                new_col = col + i
                sheet.write(row, new_col, val, style)
        elif 'dropdown' in metadata.keys():
            sheet.data_validation(row,col,row,col, {'validate': 'list','source': metadata['dropdown']})
            style = getattr(self, style_string)
            sheet.write(row, col, data, style)
        else:
            style = getattr(self, style_string)
            sheet.write(row, col, data, style)
    def fill_sheet(self, sheet, col_dict, data):
        """
        Method to fill a worksheet with simple data.
        sheet: A sheet object that has been added to a workbook.
        col_dict: A dictionary of meta-data about each column; the expected
            keys are label, width, and style.
        data: The data to be added to the sheet. Should be a container of
            containers of data, i.e. a list of lists.
        Returns an integer which is the number of the first open row at the
        bottom of the sheet.
        """
        multicol_max_length = 0
        for col, metadata in col_dict.items():
            if metadata.get('multicolumn', False) and multicol_max_length == 0:
                for datarow in data:
                    if len(datarow[col]) > multicol_max_length:
                        multicol_max_length = len(datarow[col])
        self.add_headers(sheet, col_dict, multicol_max_length)
        row = 1
        for rec in data:
            # print(rec)
            for col, metadata in col_dict.items():
                # print(col, metadata)
                # print(sheet, row, col, metadata, rec[col], multicol_max_length)
                # if rec[9]:
                #     sheet.write_comment(row,0,rec[9])
                self._write_data_to_column(sheet, row, col, metadata, rec[col], multicol_max_length)
            row += 1
        return row

    def fill_sheet_from_profile_objects(self, sheet, col_dict, object_list):
        """
        Probably going to delete this as it is too specific for FB Page objects.
        """
        colnumbers = sorted(col_dict.keys())
        self.add_headers(sheet, col_dict)
        row = 1
        for rec in object_list:
            for col in colnumbers:
                colinfo = col_dict[col]
                style = colinfo['style']
                if 'special' in colinfo:
                    if colinfo['special'] == 'constant':
                        value = colinfo['attr']
                    elif colinfo['special'] == 'count':
                        c = getattr(rec, colinfo['attr'])
                        if c:
                            value = len(c)
                        else:
                            value = 0
                    elif colinfo['special'] == 'post_info':
                        post_id = post.post_id.split('_')[1]
                        value = post_info[post_id][colinfo['attr']]
                        if isinstance(value, list):
                            value = ', '.join(value)
                    elif colinfo['special'] == 'short_url':
                        value = get_post_message_short_url(post.post_message)
                    elif colinfo['special'] == 'short_url_count':
                        short_url = get_post_message_short_url(post.post_message)
                        value = get_short_url_click_count(short_url, url_tracking)
                elif 'page' in colinfo and colinfo['page']:
                    value = getattr(page, colinfo['attr'])
                elif colinfo['style'] == date_style:
                    value = getattr(post, colinfo['attr'])
                    lmd = time.strptime(value, "%Y-%m-%dT%H:%M:%S+0000")
                    value = str(time.strftime('%m/%d/%Y', lmd))
                elif colinfo['style'] == time_style:
                    value = getattr(post, colinfo['attr'])
                    lmd = time.strptime(value, "%Y-%m-%dT%H:%M:%S+0000")
                    value = str(time.strftime('%H:%M', lmd))
                elif colinfo['attr'] == 'post_message':
                    value = getattr(post, colinfo['attr'])
                    #value = value.encode('unicode_escape').decode('utf-8')
                    value = value.replace('\n', ' ')[0:255]
                else:
                    value = getattr(post, colinfo['attr'])
                sheet.write(row, col, value, style)
            row += 1
        return row
    def add_single_row(self, sheet, row, col_dict,data):
        """
        Method to add a single row of data to a sheet. Most useful for totals.
        sheet: a sheet object that has been added to a workbook
        row: the number of the row to write the data
        col_dict: a dictionary of meta-data about each column
        data: a container of data, i.e. a list or tuple
        """
        for col, metadata in col_dict.items():
            style = getattr(self, metadata['style'])
            sheet.write(row, col, data[col], style)
        row += 1
        return row

    def add_single_row_from_list(self, sheet, row, col_dict,data):
        """
        NEw
        """
        for col, metadata in col_dict.items():
            style = getattr(self, metadata['style'])
            sheet.write(row, col, data, style)
        row += 1
        return row
    
    def add_single_row_shift(self, sheet, row, col_dict, shift,data):
        """
        Method to add a single row of data to a sheet. Most useful for totals.
        sheet: a sheet object that has been added to a workbook
        row: the number of the row to write the data
        col_dict: a dictionary of meta-data about each column
        data: a container of data, i.e. a list or tuple
        """
        for col, metadata in col_dict.items():
            style = getattr(self, metadata['style'])
            sheet.write(row, col+shift, data[col], style)
        row += 1
        return row

    def add_pandas_table(self, sheet, col_dict, df,row,shift):
        """
        TODO
        sheet: a sheet object that has been added to a workbook
        row: the number of the row to write the data
        col_dict: a dictionary of meta-data about each column
        df: a dataframe of data
        """
        (max_row, max_col) = df.shape
        column_settings = [{'header': column} for column in df.columns]
        sheet.add_table(row, shift, max_row+row, max_col - 1 +shift, {'columns': column_settings})
        sheet.set_column(shift, max_col - 1, 12)
        return max_row+row+1
    
    def close_workbook(self):
        """
        Closing and saving the workbook.
        """
        self.workbook.close()





    def add_single_row_new_way(self, sheet, row, col, col_dict,data):
        """
        NEW
        """
        # print(col_dict['style'])
        style = getattr(self, col_dict['style'])
        sheet.set_column(col, col, col_dict['width'])
        sheet.write(row, col, data, style)
        row += 1
        return row