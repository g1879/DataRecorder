# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from .base import BaseRecorder
from .functions import _data_to_list


class Filler(BaseRecorder):
    SUPPORTS = ('xlsx',)

    def __init__(self, file_path: Union[str, Path],
                 cache_size: int = 50,
                 key_cols: Union[str, int, list, tuple] = 1,
                 begin_row: int = 2,
                 sign_col: Union[str, int] = 2):
        """初始化                                  \n
        :param file_path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件
        """
        super().__init__(file_path, cache_size)
        self.key_cols = key_cols
        self.begin_row = begin_row
        self.sign_col = sign_col

    def set_file_path(self, path: Union[str, Path], key_cols: Union[str, int, list, tuple]):
        if not Path(path).exists():
            raise FileNotFoundError('文件不存在')
        self.file_path = path
        self.key_cols = key_cols

    @property
    def key_cols(self):
        return self._key_cols

    @key_cols.setter
    def key_cols(self, cols: Union[str, int, list, tuple]):
        self._key_cols = (cols,) if isinstance(cols, (int, str)) else cols

    @property
    def begin_row(self):
        return self._begin_row

    @begin_row.setter
    def begin_row(self, row: int):
        if not isinstance(row, int):
            raise TypeError('row值只能是int')
        self._begin_row = row

    @property
    def sign_col(self):
        return self._sign_col

    @sign_col.setter
    def sign_col(self, col: Union[str, int]):
        if not isinstance(col, (int, str)):
            raise TypeError('col值只能是int或str')
        self._sign_col = col

    @property
    def keys(self) -> list:
        """返回key列内容，第一位为行号，其余为key列的值  \n
        eg.[3, '名称', 'id']
        :return: 行号及key列值组成的列表
        """
        wb = load_workbook(self.file_path, read_only=True)
        ws = wb.active

        keys = []
        for row in range(self.begin_row, ws.max_row + 1):
            sign_col = self.sign_col if isinstance(self.sign_col, int) else column_index_from_string(self.sign_col)
            if ws.cell(row, sign_col).value is not None:
                continue

            key = [row]
            for col in self.key_cols:
                col = col if isinstance(col, int) else column_index_from_string(col)
                key.append(ws.cell(row, col).value)

            keys.append(key)

        wb.close()
        return keys

    def add_data(self, data: Union[list, tuple, dict]):
        """添加数据                                                 \n
        数据格式：两位的list或tuple，第一位为行号，第二位为数据，数据可以是list,tuple,dict
        :param data: 要添加的内容，第一位为行号，其余为要添加的内容
        :return:
        """
        if not data:
            return
        if isinstance(data, dict) or isinstance(data[0], int):  # 只有一个数据的情况
            data = (data,)

        new_data = []
        for d in data:
            if isinstance(d, dict):
                d = list(d.values())

            length = len(d)
            if not isinstance(d, (list, tuple, dict)) or length < 2 or not isinstance(d[0], int):
                raise ValueError('数据项必须为长度大于2的list、tuple或dict，且第一位为int代表行号。')

            if length == 2 and isinstance(d[1], (list, tuple, dict)):  # 只有两位且第二位是数据集
                d1 = list(d[1].values()) if isinstance(d[1], dict) else list(d[1:])
                d = [d[0]].extend(d1)

            new_data.append(d)

        self._data.extend(new_data)

        if len(self._data) >= self.cache_size:
            self.record()

    def record(self):
        """记录数据"""
        if not self._data:
            return

        if self.file_type == 'xlsx':
            col = self.sign_col if isinstance(self.sign_col, int) else column_index_from_string(self.sign_col)
            _fill_to_xlsx(self.file_path, self._data, self._before, self._after, col)

        self._data = []


def _fill_to_xlsx(file_path: str,
                  data: Union[list, tuple],
                  before: Union[list, dict] = None,
                  after: Union[list, dict] = None,
                  col: int = None) -> None:
    """记录数据到xlsx文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :param col: 开始记录的列号
    :return: None
    """
    wb = load_workbook(file_path)
    ws = wb.active

    for i in data:
        for key, j in enumerate(_data_to_list(i[1:], before, after)):
            ws.cell(i[0], col + key).value = j

    wb.save(file_path)
    wb.close()
