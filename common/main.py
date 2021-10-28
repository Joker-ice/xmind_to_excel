import common.util as util
from common.operation_excel import WriteExcel
from common.operation_xmind import ReadXmind
import settings.conf as c
import os
import random


class TestProgram():
    def __init__(self, index=None):
        self.index = index
        self.cwd = util.filelist.cwd_path
        self.filenames = util.filelist.get_files()
        self.isp0 = c.isp0
        self.to_excel = c.to_excel
        self.title_name = c.title_name
        self.col_value = c.col_value
        self.p0_title_name = c.p0_title_name
        self.p0_col_value = c.p0_col_value
        self.write_excel()

    def write_excel(self):
        if isinstance(self.filenames, str):
            self.common_req(self.filenames, self.isp0, self.to_excel)
        elif isinstance(self.filenames, list):
            for file in self.filenames:
                # file_path = self.cwd + r'\%s' % file
                self.common_req(file, self.isp0, self.to_excel)

    def common_req(self, file_path, isp0, to_excel):
        readxmind = ReadXmind(self.index, file_path)
        sheets = readxmind.all_sheet
        new_sheets, p0_sheets = self.get_format(sheets, isp0, to_excel)
        first_title = readxmind.first_title
        if to_excel:
            write_path = self.cwd + r'\%s.xlsx' % first_title
            WriteExcel(new_sheets, self.title_name, write_path)
        if isp0:
            # print(isp0)
            write_path = self.cwd + r'\%s_p0.xlsx' % first_title
            WriteExcel(p0_sheets, self.p0_title_name, write_path)

    def get_format(self, topics, isp0, to_excel):
        root, p0_root = [], []
        if to_excel:
            for i in topics:
                num = 0
                if i[-1].strip().lower() != 'p0':
                    i.append('p1')
                new_col_value = self.col_value.copy()
                for value in self.col_value:
                    if isinstance(value, int):
                        new_col_value[num] = i[value]
                    elif isinstance(value, list):
                        s = ''
                        if value == ['step']:
                            v_num = 1
                            for new_data in i[1:]:
                                if v_num == len(i) - 2:
                                    break
                                s += '%s.%s\n' % (str(v_num), new_data)
                                v_num += 1
                        else:
                            for v in value:
                                if isinstance(v, int):
                                    s += i[v]
                                elif isinstance(v, str):
                                    try:
                                        s += str(eval(v))
                                    except Exception as e:
                                        s += v
                        new_col_value[num] = s
                    num += 1
                root.append(new_col_value)
        if isp0:
            key = 1
            for i in topics:
                if i[-1].strip().lower() == 'p0':
                    num = 0
                    new_col_value = self.p0_col_value.copy()
                    for value in self.p0_col_value:
                        if isinstance(value, int):
                            new_col_value[num] = i[value]
                        elif isinstance(value, list):
                            s = ''
                            if value == ['step']:
                                v_num = 1
                                for new_data in i[1:]:
                                    if v_num == len(i) - 2:
                                        break
                                    s += '%s.%s\n' % (str(v_num), new_data)
                                    # print(s)
                                    v_num += 1
                            elif value == ['index']:
                                s = str(key)
                                key += 1
                            else:
                                for v in value:
                                    if isinstance(v, int):
                                        s += i[v]
                                    elif isinstance(v, str):
                                        s += v
                            new_col_value[num] = s
                        num += 1
                    p0_root.append(new_col_value)
        return root, p0_root


main = TestProgram

# if __name__ == '__main__':
#     main = TestProgram()
