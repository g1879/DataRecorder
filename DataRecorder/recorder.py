# -*- coding:utf-8 -*-
from pathlib import Path
from time import sleep
from typing import Union

from openpyxl import Workbook, load_workbook

from .base import BaseRecorder
from .cell_style import CellStyle
from .setter import RecorderSetter
from .tools import ok_list


class Recorder(BaseRecorder):
    """用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销。
    退出时能自动记录数据（xlsx格式除外），避免因异常丢失。
    """
    SUPPORTS = ('any',)

    def __init__(self, path=None, cache_size=None):
        super().__init__(path=path, cache_size=cache_size)
        self._follow_styles = False
        self._row_styles = None
        self._row_styles_len = None

        self._col_height = None
        self._style = None

    @property
    def set(self):
        """返回用于设置属性的对象"""
        if self._setter is None:
            self._setter = RecorderSetter(self)
        return self._setter

    def add_data(self, data):
        """添加数据，可一次添加多条数据
        :param data: 插入的数据，元组或列表
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
            self._data.append(data)
        else:  # 二维数组
            self._data.extend(data)

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def _record(self):
        """记录数据"""
        if self.type == 'xlsx':
            self._to_xlsx()
        elif self.type == 'csv':
            self._to_csv()
        elif self.type == 'json':
            self._to_json()
        else:
            self._to_txt()

    def _to_xlsx(self):
        """记录数据到xlsx文件"""
        if Path(self.path).exists():
            wb = load_workbook(self.path)
            if self.table:
                ws = wb[self.table] if self.table in [i.title for i in wb.worksheets] else wb.create_sheet(
                    title=self.table)
            else:
                ws = wb.active

            if self._follow_styles and self._row_styles is None:
                row_num = ws.max_row
                self._row_styles = [CellStyle(i) for i in ws[row_num]]
                self._row_styles_len = len(self._row_styles)
                self._col_height = ws.row_dimensions[row_num].height

        else:
            if self._col_height or self._row_styles:
                wb = Workbook(write_only=False)
                ws = wb.active
                if self.table:
                    ws.title = self.table
            else:
                wb = Workbook(write_only=True)
                ws = wb.create_sheet(title=self.table)

            title = _get_title(self._data[0], self._before, self._after)
            if title is not None:
                ws.append(ok_list(title, True))

        for i in self._data:
            data = ok_list(self._data_to_list(i), True)
            ws.append(data)

            if self._col_height is not None:
                ws.row_dimensions[ws.max_row].height = self._col_height

            if self._row_styles:
                groups = zip(ws[ws.max_row], self._row_styles)
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
            all_data = [' '.join(ok_list(self._data_to_list_or_dict(i), as_str=True)) for i in self._data]
            f.write('\n'.join(all_data) + '\n')

    def _to_json(self):
        """记录数据到json文件"""
        from json import load, dump
        if Path(self.path).exists():
            with open(self.path, 'r', encoding=self.encoding) as f:
                json_data = load(f)

            for i in self._data:
                json_data.append(ok_list(self._data_to_list_or_dict(i)))

        else:
            json_data = [ok_list(self._data_to_list_or_dict(i)) for i in self._data]

        with open(self.path, 'w', encoding=self.encoding) as f:
            dump(json_data, f)

    def _data_to_list_or_dict(self, data):
        """将传入的数据转换为列表或字典形式，添加前后列数据，用于记录到txt或json
        :param data: 要处理的数据
        :return: 转变成列表或字典形式的数据
        """
        if isinstance(data, (list, tuple)):
            return self._data_to_list(data)

        elif isinstance(data, dict):
            if isinstance(self.before, dict):
                data = {**self.before, **data}

            if isinstance(self.after, dict):
                data = {**data, **self.after}

            return data


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
