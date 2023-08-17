# -*- coding:utf-8 -*-
from pathlib import Path
from time import sleep
from typing import Union

from openpyxl import Workbook, load_workbook

from .base import BaseRecorder
from .style.cell_style import CellStyleCopier
from .setter import RecorderSetter
from .tools import ok_list, data_to_list_or_dict


class Recorder(BaseRecorder):
    SUPPORTS = ('csv', 'xlsx', 'json', 'txt')

    def __init__(self, path=None, cache_size=None):
        """用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        super().__init__(path=path, cache_size=cache_size)
        self._delimiter = ','  # csv文件分隔符
        self._quote_char = '"'  # csv文件引用符
        self._follow_styles = False
        self._col_height = None
        self._style = None

    @property
    def set(self):
        """返回用于设置属性的对象"""
        if self._setter is None:
            self._setter = RecorderSetter(self)
        return self._setter

    @property
    def delimiter(self):
        """返回csv文件分隔符"""
        return self._delimiter

    @property
    def quote_char(self):
        """返回csv文件引用符"""
        return self._quote_char

    def add_data(self, data, table=None):
        """添加数据，可一次添加多条数据
        :param data: 插入的数据，任意格式
        :param table: 要写入的数据表，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.1)

        if not isinstance(data, (list, tuple, dict)):
            data = (data,)
        if not data:
            data = (None,)

        # 一维数组
        if (isinstance(data, (list, tuple)) and not isinstance(data[0], (list, tuple, dict))) or isinstance(data, dict):
            data = [data]
            self._data_count += 1
        else:  # 二维数组
            self._data_count += len(data)

        if self._type != 'xlsx':
            self._data.extend(data)

        else:
            if table is None:
                table = self._table
            elif isinstance(table, bool):
                table = None

            self._data.setdefault(table, []).extend(data)

        if 0 < self.cache_size <= self._data_count:
            self.record()

    def _record(self):
        """记录数据"""
        if self.type == 'csv':
            self._to_csv()
        elif self.type == 'xlsx':
            self._to_xlsx()
        elif self.type == 'json':
            self._to_json()
        elif self.type == 'txt':
            self._to_txt()

    def _to_xlsx(self):
        """记录数据到xlsx文件"""
        if Path(self.path).exists():
            new_file = False
            wb = load_workbook(self.path)

        else:
            new_file = True
            if self._col_height or self._follow_styles or self._style or len(self._data) > 1:
                wb = Workbook(write_only=False)
            else:
                wb = Workbook(write_only=True)
                wb.create_sheet('Sheet1')

        tables = [i.title for i in wb.worksheets]
        for table, data in self._data.items():
            _row_styles = None
            _col_height = None
            new_sheet = False

            if table is None:
                ws = wb.active

            elif table in tables:
                ws = wb[table]

            elif new_file:
                ws = wb.active
                tables.remove(ws.title)
                ws.title = table
                tables.append(table)
                new_sheet = True

            else:
                ws = wb.create_sheet(title=table)
                tables.append(table)
                new_sheet = True

            if new_file or new_sheet:
                new_file = False
                _add_title(ws, data[0], self._before, self._after)

            elif self._follow_styles:
                row_num = ws.max_row
                _row_styles = [CellStyleCopier(i) for i in ws[row_num]]
                _col_height = ws.row_dimensions[row_num].height

            # ====================================

            for i in data:
                d = ok_list(self._data_to_list(i), True)
                ws.append(d)

                if _col_height is not None:
                    ws.row_dimensions[ws.max_row].height = self._col_height

                if _row_styles:
                    groups = zip(ws[ws.max_row], _row_styles)
                    for g in groups:
                        g[1].to_cell(g[0])

                elif self._style:
                    for c in ws[ws.max_row]:
                        self._style.to_cell(c)

        wb.save(self.path)
        wb.close()

    def _to_csv(self):
        """记录数据到csv文件"""
        from csv import writer
        title = _get_title(self._data[0], self._before, self._after) if not Path(self.path).exists() else None

        with open(self.path, 'a+', newline='', encoding=self.encoding) as f:
            csv_write = writer(f, delimiter=self.delimiter, quotechar=self.quote_char)
            if title:
                csv_write.writerow(ok_list(title))
            for i in self._data:
                csv_write.writerow(ok_list(self._data_to_list(i)))

    def _to_txt(self):
        """记录数据到txt文件"""
        with open(self.path, 'a+', encoding=self.encoding) as f:
            all_data = [' '.join(ok_list(data_to_list_or_dict(self, i), as_str=True)) for i in self._data]
            f.write('\n'.join(all_data) + '\n')

    def _to_json(self):
        """记录数据到json文件"""
        from json import load, dump
        if Path(self.path).exists():
            with open(self.path, 'r', encoding=self.encoding) as f:
                json_data = load(f)

            for i in self._data:
                json_data.append(ok_list(data_to_list_or_dict(self, i)))

        else:
            json_data = [ok_list(data_to_list_or_dict(self, i)) for i in self._data]

        with open(self.path, 'w', encoding=self.encoding) as f:
            dump(json_data, f)


def _add_title(ws, data, before, after):
    """向空sheet添加表头"""
    title = _get_title(data, before, after)
    if title is not None:
        ws.append(ok_list(title, True))


def _get_title(data: Union[list, dict],
               before: Union[list, dict, None] = None,
               after: Union[list, dict, None] = None) -> Union[list, None]:
    """获取表头列表
    :param data: 数据列表或字典
    :param before: 数据前的列
    :param after: 数据后的列
    :return: 表头列表
    """
    if isinstance(data, (tuple, list)):
        return None

    return_list = []
    for i in (before, data, after):
        if isinstance(i, dict):
            return_list.extend(list(i))
        elif i is None:
            pass
        elif isinstance(i, list):
            return_list.extend(['' for _ in range(len(i))])
        else:
            return_list.extend([''])

    return return_list
