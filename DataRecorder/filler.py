# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, List

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from .base import BaseRecorder
from .functions import _data_to_list


class Filler(BaseRecorder):
    """Filler类用于根据现有文件中的关键字向文件填充数据"""
    SUPPORTS = ('xlsx', 'csv')

    def __init__(self,
                 path: Union[str, Path],
                 cache_size: int = 50,
                 key_cols: Union[str, int, list, tuple] = 1,
                 begin_row: int = 2,
                 sign_col: Union[str, int] = 2,
                 sign: str = None,
                 data_col: int = None):
        """初始化                                            \n
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件
        :param key_cols: 作为关键字的列，可以是多列，从1开始
        :param begin_row: 数据开始的行，默认表头一行
        :param sign_col: 用于判断是否已填数据的列，从1开始
        :param sign: 按这个值判断是否已填数据
        :param data_col: 要填入数据的第一列，从1开始
        """
        super().__init__(path, cache_size)
        self.key_cols = key_cols
        self.begin_row = begin_row
        self.sign_col = sign_col
        self.sign = sign
        self.data_col = data_col or sign_col

    def set_path(self,
                 path: Union[str, Path],
                 key_cols: Union[str, int, list, tuple] = 1,
                 begin_row: int = None,
                 sign_col: Union[str, int] = None,
                 sign: str = None,
                 data_col: int = None) -> None:
        """设置文件路径                             \n
        :param path: 保存的文件路径
        :param key_cols: 作为关键字的列，可以是多列
        :param begin_row: 数据开始的行，默认表头一行
        :param sign_col: 用于判断是否已填数据的列
        :param sign: 按这个值判断是否已填数据
        :param data_col: 要填入数据的第一列
        """
        if not Path(path).exists():
            raise FileNotFoundError('文件不存在')
        self.begin_row = begin_row or self.begin_row
        self.sign_col = sign_col or self.sign_col
        self.sign = sign or self.sign
        self.data_col = sign or self.data_col
        self.__init__(path, self.cache_size, key_cols, begin_row, sign_col, sign, data_col)

    @property
    def key_cols(self) -> Union[list, tuple]:
        """返回作为关键字的列或列的集合"""
        return self._key_cols

    @key_cols.setter
    def key_cols(self, cols: Union[str, int, list, tuple]) -> None:
        """设置作为关键字的列，可以是多列                    \n
        :param cols: 列号或列名，或它们组成的list或tuple
        :return: None
        """
        self._key_cols = (cols,) if isinstance(cols, (int, str)) else cols

    @property
    def begin_row(self) -> Union[str, int]:
        """返回数据开始的行号，从1开始"""
        return self._begin_row

    @begin_row.setter
    def begin_row(self, row: int) -> None:
        """设置数据开始的行
        :param row: 行号
        :return: None
        """
        if not isinstance(row, int) or row < 1:
            raise TypeError('row值只能是int，且必须大于0')
        self._begin_row = row

    @property
    def sign_col(self) -> Union[str, int]:
        """返回用于判断是否已填数据的列"""
        return self._sign_col

    @sign_col.setter
    def sign_col(self, col: Union[str, int]) -> None:
        """设置用于判断是否已填数据的列
        :param col: 列号或列名
        :return: None
        """
        if not isinstance(col, (int, str)):
            raise TypeError('col值只能是int或str')
        self._sign_col = col

    @property
    def keys(self) -> list:
        """返回key列内容，第一位为行号，其余为key列的值  \n
        eg.[3, '名称', 'id']
        """
        if self.type in ('xlsx', 'csv'):
            return _get_keys(self.path, self.begin_row, self.sign_col, self.sign, self.key_cols)

    def add_data(self, data: Union[list, tuple, dict]) -> None:
        """添加数据                                                                   \n
        数据格式：两位的list或tuple，第一位为行号，第二位为数据，数据可以是list, tuple, dict
        :param data: 要添加的内容，第一位为行号，其余为要添加的内容
        :return: None
        """
        if not data:
            return
        if isinstance(data, dict) or isinstance(data[0], int):  # 只有一个数据的情况
            data = (data,)

        new_data = []
        for item in data:
            if isinstance(item, dict):
                item = list(item.values())

            length = len(item)
            if not isinstance(item, (list, tuple, dict)) or length < 2 or not isinstance(item[0], int):
                raise ValueError('数据项必须为长度大于2的list、tuple或dict，且第一位为int代表行号。')

            if length == 2 and isinstance(item[1], (list, tuple, dict)):  # 只有两位且第二位是数据集
                if isinstance(item[1], dict):
                    vals = list(item[1].values())
                elif isinstance(item[1], tuple):
                    vals = list(item[1])
                else:
                    vals = item[1]
                item = [item[0]]
                item.extend(vals)

            new_data.append(item)

        self._data.extend(new_data)

        if len(self._data) >= self.cache_size:
            self.record()

    def record(self) -> None:
        """记录数据"""
        if not self._data:
            return

        col = self.data_col if isinstance(self.data_col, int) else column_index_from_string(self.data_col)
        if self.type == 'xlsx':
            _fill_to_xlsx(self.path, self._data, self._before, self._after, col)
        elif self.type == 'csv':
            _fill_to_csv(self.path, self._data, self._before, self._after, col)

        self._data = []

    def fill(self, func, *args) -> None:
        """接收一个方法，根据keys自动填充数据。每条key调用一次该方法，并根据方法返回的内容进行填充  \n
        方法第一个参数必须是keys，用于接收关键字列                                            \n
        :param func: 用于获取数据的方法，返回要填充的数据
        :param args: 该方法的参数
        :return: None
        """
        for i in self.keys:
            print(i)
            res = [i[0], func(i[1:], *args)]
            self.add_data(res)
        self.record()


def _get_keys(path: str,
              begin_row: int,
              sign_col: int,
              sign: str,
              key_cols: Union[list, tuple]) -> List[list]:
    """返回key列内容，第一位为行号，其余为key列的值       \n
    eg.[3, '名称', 'id']
    :param path: 文件路径
    :param begin_row: 数据起始行
    :param sign_col: 用于判断是否已填数据的列，从1开始
    :param sign: 按这个值判断是否已填数据
    :param key_cols: 关键字所在列，可以是多列
    :return: 关键字组成的列表
    """
    key_cols = list(map(lambda x: x - 1 if isinstance(x, int) else column_index_from_string(x) - 1, key_cols))
    sign_col -= 1
    begin_row = begin_row or 1

    if path.endswith('xlsx'):
        df = pd.read_excel(path, header=None, skiprows=begin_row - 1)
    elif path.endswith('csv'):
        df = pd.read_csv(path, header=None, skiprows=begin_row - 1)
    else:
        raise TypeError('只支持xlsx和csv格式。')

    if sign_col <= df.shape[1]:
        df = df[df[sign_col].isna()] if sign is None else df[df[sign_col] == sign]
    df = df[key_cols]
    df.index += begin_row
    df = df.where(df.notnull(), None)
    return [list(i) for i in df.itertuples()]


def _fill_to_xlsx(file_path: str,
                  data: Union[list, tuple],
                  before: Union[list, dict] = None,
                  after: Union[list, dict] = None,
                  col: int = None) -> None:
    """填写数据到xlsx文件            \n
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


def _fill_to_csv(file_path: str,
                 data: Union[list, tuple],
                 before: Union[list, dict] = None,
                 after: Union[list, dict] = None,
                 col: int = None) -> None:
    """填写数据到xlsx文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :param col: 开始记录的列号
    :return: None
    """
    df = pd.read_csv(file_path, header=None)
    df = df.where(df.notnull(), None)
    df_width = df.shape[1]
    full_width = col + len(data[0])

    for i in range(full_width - df_width - 2):
        df[df_width + i] = None

    for i in data:
        for k, j in enumerate(_data_to_list(i[1:], before, after)):
            df.loc[i[0] - 1, col + k - 1] = j

    df.to_csv(file_path, header=False, index=False)
