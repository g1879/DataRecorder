# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, List
from csv import reader as csv_reader, writer as csv_writer

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

from .base import BaseRecorder, _data_to_list


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
                 data_col: Union[int, str] = None):
        """初始化                                                  \n
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，传入0表示不自动保存
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
        if self.type == 'csv':
            return _get_csv_keys(self.path, self.begin_row, self.sign_col, self.sign, self.key_cols, self.encoding)
        elif self.type == 'xlsx':
            return _get_xlsx_keys(self.path, self.begin_row, self.sign_col, self.sign, self.key_cols)

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

    def add_data(self, data: Union[list, tuple, dict]) -> None:
        """添加数据                                                                   \n
        数据格式：第一位为行号或坐标（int或str），第二位开始为数据，数据可以是list, tuple, dict
        :param data: 要添加的内容，第一位为行号，其余为要添加的内容
        :return: None
        """
        if not data:
            return
        if isinstance(data, dict) or isinstance(data[0], (int, str)):  # 只有一个数据的情况
            data = (data,)

        new_data = []
        for item in data:
            if isinstance(item, dict):
                item = list(item.values())

            length = len(item)
            if not ((isinstance(item, (list, tuple, dict)) and length >= 2) or not isinstance(item[0], (int, str))):
                raise ValueError('数据项必须为长度不少于2的list、tuple或dict，且第一位为int(行号)、str(eg."B3"或"3,2")。')

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

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def record(self) -> None:
        """记录数据"""
        if not self._data:
            return

        col = self.data_col if isinstance(self.data_col, int) else column_index_from_string(self.data_col)

        while True:
            try:
                if self.type == 'xlsx':
                    _fill_to_xlsx(self.path, self._data, self._before, self._after, col)
                elif self.type == 'csv':
                    _fill_to_csv(self.path, self._data, self._before, self._after, col, self.encoding)
                break

            except PermissionError:
                input('文件被打开，保存失败，请关闭后按回车重试。')

        self._data = []

    def fill(self, func, *args) -> None:
        """接收一个方法，根据keys自动填充数据。每条key调用一次该方法，并根据方法返回的内容进行填充  \n
        方法第一个参数必须是keys，用于接收关键字列，返回的第一位必须是行号或坐标                 \n
        :param func: 用于获取数据的方法，返回要填充的数据
        :param args: 该方法的参数
        :return: None
        """
        for i in self.keys:
            print(i)
            res = [i[0], func(i, *args)]
            self.add_data(res)
        self.record()


def _get_xlsx_keys(path: str,
                   begin_row: int,
                   sign_col: Union[int, str],
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
    key_cols = [i if isinstance(i, int) else column_index_from_string(i) for i in key_cols]
    sign_col = column_index_from_string(sign_col) if isinstance(sign_col, str) else sign_col

    wb = load_workbook(path, read_only=True)
    ws = wb.active

    res_keys = []
    for row in range(begin_row, ws.max_row + 1):
        if ws.cell(row, sign_col).value == sign:
            key = [row]
            res_keys.append(key + [ws.cell(row, i).value for i in key_cols])

    wb.close()
    return res_keys


def _get_csv_keys(path: str,
                  begin_row: int,
                  sign_col: Union[int, str],
                  sign: str,
                  key_cols: Union[list, tuple],
                  encoding: str) -> List[list]:
    """返回key列内容，第一位为行号，其余为key列的值       \n
    eg.[3, '名称', 'id']
    :param path: 文件路径
    :param begin_row: 数据起始行
    :param sign_col: 用于判断是否已填数据的列，从1开始
    :param sign: 按这个值判断是否已填数据
    :param key_cols: 关键字所在列，可以是多列
    :param encoding: 字符编码，用于csv
    :return: 关键字组成的列表
    """
    key_cols = [i - 1 if isinstance(i, int) else column_index_from_string(i) - 1 for i in key_cols]
    sign_col = column_index_from_string(sign_col) if isinstance(sign_col, str) else sign_col
    sign = '' if sign is None else str(sign)
    sign_col -= 1
    begin_row -= 1
    res_keys = []

    with open(path, 'r', encoding=encoding) as f:
        reader = csv_reader(f)
        lines = list(reader)[begin_row:]

        for k, line in enumerate(lines):
            row_sign = None if sign_col > len(line) - 1 else line[sign_col]
            if row_sign == sign:
                key = [k + 1]
                res_keys.append(key + [i for num, i in enumerate(line) if num in key_cols])

    return res_keys


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
            if isinstance(i[0], int):  # 行号
                ws.cell(i[0], col + key).value = j
            elif isinstance(i[0], str):
                if ',' in i[0]:  # 坐标 如'3,2'
                    row, col = i[0].split(',')
                    ws.cell(int(row), int(col)).value = j
                else:  # 坐标 如'A8'
                    ws[i[0]].value = j
            else:
                raise TypeError(f'数据第一位必须是int（行号）、str（eg."B3"或"3,2"）。现在是：{i[0]}')

    wb.save(file_path)
    wb.close()


def _fill_to_csv(file_path: str,
                 data: Union[list, tuple],
                 before: Union[list, dict] = None,
                 after: Union[list, dict] = None,
                 col: int = None,
                 encoding: str = 'utf-8') -> None:
    """填写数据到xlsx文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :param col: 开始记录的列号
    :param encoding: 字符编码
    :return: None
    """
    with open(file_path, 'r', encoding=encoding) as f:
        reader = csv_reader(f)
        lines = list(reader)
        lines_len = len(lines)

        for i in data:
            if isinstance(i[0], int):  # 行号
                row = i[0]
            elif isinstance(i[0], str):  # 坐标 如'A8'
                if ',' in i[0]:  # 坐标 如'3,2'
                    xy = i[0].split(',')
                    row = int(xy[0])
                    col = int(xy[1])
                else:  # 坐标 如'A8'
                    xy = coordinate_from_string(i[0])
                    row = xy[1]
                    col = column_index_from_string(xy[0])
            else:
                raise TypeError(f'数据第一位必须是int（行号）、str（eg."B3"或"3,2"）。现在是：{i[0]}')

            # 若行数不够，填充行数
            for _ in range(row - lines_len):
                lines.append([])
                lines_len += 1

            # 若列数不够，填充空列
            lines[row - 1].extend([None] * (col - len(lines[row - 1]) + len(i) - 2))

            # 填充数据
            for k, j in enumerate(_data_to_list(i[1:], before, after)):
                lines[row - 1][col + k - 1] = j

        writer = csv_writer(open(file_path, 'w', encoding=encoding, newline=''))
        writer.writerows(lines)
