# -*- coding:utf-8 -*-

from tempfile import NamedTemporaryFile
import xlsxwriter

class Excel(object):

    def __init__(self, title, sheet_name='sheet1', suffix='.xlsx', title_position=0):
        """
        :param title:           标题
        :param suffix:          文件后缀名
        :param title_position   标题所在第几行
        :param sheet            当前的sheet
        :return:
        """
        self.title = title
        self.title_position = title_position
        self.row_index = title_position + 1
        self.xlsx_file = NamedTemporaryFile(suffix=suffix)
        self.workbook = xlsxwriter.Workbook(self.xlsx_file)
        self.sheets = { sheet_name: self.workbook.add_worksheet(sheet_name) }
        self.sheet = self.sheets[sheet_name]
        self._init_title(sheets=self.sheets['sheet1'], title=self.title, title_position=self.title_position)

    @classmethod
    def _init_title(cls, sheets, title, title_position, sheet_name='sheet1'):
        for index, t in enumerate(title):
            sheets.write(title_position, index, t)

    def init_title(self, sheet_name='sheet1'):
        self._init_title(sheets=self.sheet, title=self.title, title_position=self.title_position)

    def append_title(self, title):
        """
        :param title:             扩展的标题
        :return:
        """
        self.title.extend(title)
        self.init_title()

    def add_row(self, add_row_data):
        row = [''] * len(self.title)
        for title in add_row_data:
            if not title in self.title:
                continue
            title_index = self.title.index(title)
            row[title_index] = add_row_data[title]

        for index, content in enumerate(row):
            if not isinstance(content, basestring):
                content = str(content)
            self.sheet.write(self.row_index, index, content)
        self.row_index += 1