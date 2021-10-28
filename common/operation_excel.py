import logging

import pandas as pd
import numpy as np


class WriteExcel():
    def __init__(self, data=None, title_name=None, filepath=None):
        self.data = self.manage_data(data)
        self.title_name = title_name
        self.output_to_excel(filepath)

    def manage_data(self, data):
        data_list = []
        for root in data:
            root = tuple(root)
            data_list.append(root)
        # print(data_list)
        return data_list

    def handle_pd(self):
        # print(self.data)
        return pd.DataFrame(data=self.data, columns=self.title_name)

    def output_to_excel(self, filepath):
        try:
            df = self.handle_pd()
            with pd.ExcelWriter(filepath, 'xlsxwriter') as writer:
                df.to_excel(writer, 'sheet1', index=False)
                workbook = writer.book
                worksheet = writer.sheets['sheet1']
                header_format, text_format, custom_format = self.set_format(workbook)
                self.set_sheet_format(worksheet, df, header_format, text_format, custom_format)
                self.set_width(df, worksheet)
                print('>>>>>>>To successfully generate the file: %s' % filepath.split('\\')[-1])
        except Exception as e:
            raise e

    def set_format(self, workbook):
        header_format = workbook.add_format({
            'bold': True, # 字体加粗
            'text_wrap': True, # 是否自动换行
            'valign': 'vcenter',  #垂直对齐方式
            'align': 'center', # 水平对齐方式
            'font': '宋体',
            'font_size': 14,
            # 'fg_color': '#D7E4BC', # 单元格背景颜色
            # 'border': 2  # 单元格边框宽度
        })
        text_format = workbook.add_format({
            'bold': False, # 字体加粗
            'text_wrap': True, # 是否自动换行
            'valign': 'vcenter',  #垂直对齐方式
            'align': 'center', # 水平对齐方式
            'font': '宋体',
            'font_size': 11,
            # 'fg_color': '#D7E4BC', # 单元格背景颜色
            # 'border': 2  # 单元格边框宽度
        })
        custom_format = workbook.add_format({
            'bold': False, # 字体加粗
            'text_wrap': True, # 是否自动换行
            'valign': 'vcenter',  #垂直对齐方式
            'align': 'left', # 水平对齐方式
            'font': '宋体',
            'font_size': 11,
            # 'fg_color': '#D7E4BC', # 单元格背景颜色
            # 'border': 2  # 单元格边框宽度
        })
        return header_format, text_format, custom_format

    def set_sheet_format(self, worksheet, df, header_format, text_format, custom_format):
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        # Write the row with the defined format.
        for index, value in df.iterrows():
            # print(index, " -- > ", value.values)
            colnum = 0
            for v in value:
                # print(v)
                if '\n' in v:
                    worksheet.write(index+1, colnum, v, custom_format)
                else:
                    worksheet.write(index+1, colnum, v, text_format)
                colnum += 1

    def set_width(self, df, worksheet):
        # 计算表头的字符宽度
        column_widths = (
            df.columns.to_series()
              .apply(lambda x: len(x.encode('utf8'))).values
        )
        #  计算每列的最大字符宽度
        max_widths = (
            df.astype(str)
              .applymap(lambda x: len(x.encode('utf8')))
              .agg(max).values
        )
        # print(max_widths)
        # 计算整体最大宽度
        widths = np.max([column_widths, max_widths], axis=0)
        for i, width in enumerate(widths):
            if width > 100:
                worksheet.set_column(i, i, 57)
            elif width < 10:
                worksheet.set_column(i, i, 10)
            else:
                worksheet.set_column(i, i, width)


if __name__ == '__main__':
    r = WriteExcel()
    # dp = pandas.DataFrame(data=data_list, columns=title_list)
    # dp.to_excel(cwd, 'sheet1', index=False)
