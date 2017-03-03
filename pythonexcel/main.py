# -*- coding:utf-8 -*-

from tempfile import NamedTemporaryFile
import xlsxwriter

class Excel(object):

    def __init__(self, **kwargs):
        """
        :param title:           标题
        :param suffix:          文件后缀名
        :param title_position   标题所在第几行
        :param sheet            当前的sheet
        :return:
        """
        self.sheets = {}
        suffix = kwargs.get('suffix', '.xlsx')
        self.xlsx_file = NamedTemporaryFile(suffix=suffix)
        self.workbook = xlsxwriter.Workbook(self.xlsx_file)
        self._to_sheet(**kwargs)
        self._init_title(sheets=self.sheets[self.sheet_name]['worksheet'], title=self.title, title_position=self.title_position)

    @property
    def title(self):
        """ 获得当前 sheet 的 标题
        """
        return self.sheets[self.sheet_name]['title']

    @property
    def title_position(self):
        """ 获得当前 sheet 的 标题位置
        """
        return self.sheets[self.sheet_name]['title_position']

    @property
    def sheet(self):
        """ 获得当前 sheet 实例
        """
        return self.sheets[self.sheet_name]['worksheet']

    @property
    def row_index(self):
        return self.sheets[self.sheet_name]['row_index']

    def _to_sheet(self, **kwargs):
        """ 切换到指定 sheet
        """
        self.sheet_name = kwargs.get('sheet_name', 'sheet1')
        if not self.sheet_name in self.sheets:
            self.sheets[self.sheet_name] = {}
            self.sheets[self.sheet_name]['title'] = kwargs.get('title', '')
            self.sheets[self.sheet_name]['title_position'] = kwargs.get('title_position', 0)
            self.sheets[self.sheet_name]['row_index'] = self.title_position + 1
            self.sheets[self.sheet_name]['worksheet'] = self.workbook.add_worksheet(self.sheet_name)

    @classmethod
    def _init_title(cls, sheets, title, title_position):
        for index, t in enumerate(title):
            sheets.write(title_position, index, t)

    def init_title(self, title=None):
        if not title is None:
            if len(self.sheets[self.sheet_name]['title']) > len(title):
                title += [u'　'] * (len(self.sheets[self.sheet_name]['title']) - len(title))
            self.sheets[self.sheet_name]['title'] = title
        self._init_title(sheets=self.sheet, title=self.title, title_position=self.title_position)

    def append_title(self, title):
        """
        :param title:             扩展的标题
        :return:
        """
        self.sheets[self.sheet_name]['title'].extend(title)
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
        self.sheets[self.sheet_name]['row_index'] += 1

    def to_sheet(self, sheet_name):
        """ go to another sheet
        """
        if not sheet_name in self.sheets:
            self._to_sheet(sheet_name=sheet_name)
        self.sheet_name = sheet_name