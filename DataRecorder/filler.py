# -*- coding:utf-8 -*-
from csv import reader as csv_reader, writer as csv_writer
from pathlib import Path
from typing import Union, List

from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string

from .base import BaseRecorder, _data_to_list, _parse_coord, _process_content


class Filler(BaseRecorder):
    """Filler类用于根据现有文件中的关键字向文件填充数据"""

    def __init__(self,
                 path: Union[str, Path],
                 cache_size: int = None,
                 key_cols: Union[str, int, list, tuple] = 1,
                 begin_row: int = 2,
                 sign_col: Union[str, int] = 2,
                 sign: str = None,
                 data_col: Union[int, str] = None) -> None:
        """初始化                                                    \n
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
            return _get_csv_keys(self.path, self.begin_row, self.sign_col, self.sign, self.key_cols, self.encoding,
                                 self.delimiter, self.quote_char)
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

    def add_data(self,
                 data: Union[list, tuple, dict, int, str, float],
                 coord: Union[list, tuple, str, int] = None) -> None:
        """添加数据                                                                   \n
        数据格式：第一位为行号或坐标（int或str），第二位开始为数据，数据可以是list, tuple, dict
        :param data: 要添加的内容，第一位为行号，其余为要添加的内容
        :param coord: 要添加数据的坐标，仅用于添加一行数据
        :return: None
        """
        if not data:
            return

        if coord is not None:
            coord = _parse_coord(coord, self.data_col)
            coord = f'{coord[0]},{coord[1]}'
            data = (coord, data)

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
                    val = list(item[1].values())
                elif isinstance(item[1], tuple):
                    val = list(item[1])
                else:
                    val = item[1]
                item = [item[0]]
                item.extend(val)

            new_data.append(item)

        self._data.extend(new_data)

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def set_link(self,
                 coord: Union[int, str, tuple, list],
                 link: str,
                 content: Union[int, str, float] = None) -> None:
        """为单元格设置超链接                          \n
        :param coord: 单元格坐标
        :param link: 超链接
        :param content: 单元格内容
        :return: None
        """
        if self.type != 'xlsx':
            raise TypeError('set_link()只支持xlsx格式。')

        self.add_data(('set_link', coord, link, content))

    def _record(self) -> None:
        """记录数据"""
        if self.type == 'xlsx':
            self._to_xlsx()
        elif self.type == 'csv':
            self._to_csv()

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

    def _to_xlsx(self) -> None:
        """填写数据到xlsx文件"""
        wb = load_workbook(self.path) if Path(self.path).exists() else Workbook()
        ws = wb.active
        col = self.data_col if isinstance(self.data_col, int) else column_index_from_string(self.data_col)

        for i in self._data:
            if i[0] == 'set_link':
                row, col = _parse_coord(i[1], col)
                cell = ws.cell(row, col)
                cell.hyperlink = i[2]
                if i[3] is not None:
                    cell.value = i[3]
            else:
                row, col = _parse_coord(i[0], col, (list, tuple))
                for key, j in enumerate(_data_to_list(i[1:], self._before, self._after)):
                    ws.cell(row, col + key).value = _process_content(j)

        wb.save(self.path)
        wb.close()

    def _to_csv(self) -> None:
        """填写数据到xlsx文件"""
        col = self.data_col if isinstance(self.data_col, int) else column_index_from_string(self.data_col)

        if not Path(self.path).exists():
            with open(self.path, 'w', encoding=self.encoding):
                pass

        with open(self.path, 'r', encoding=self.encoding) as f:
            reader = csv_reader(f, delimiter=self.delimiter, quotechar=self.quote_char)
            lines = list(reader)
            lines_len = len(lines)

            for i in self._data:
                now_data = _data_to_list(i[1:], self._before, self._after)
                row, col = _parse_coord(i[0], col, (list, tuple))

                # 若行数不够，填充行数
                for _ in range(row - lines_len):
                    lines.append([])
                    lines_len += 1

                row_num = row - 1
                # 若列数不够，填充空列
                lines[row_num].extend([None] * (col - len(lines[row_num]) + len(now_data) - 1))

                # 填充数据
                for k, j in enumerate(now_data):
                    lines[row_num][col + k - 1] = _process_content(j)

            writer = csv_writer(open(self.path, 'w', encoding=self.encoding, newline=''), delimiter=self.delimiter,
                                quotechar=self.quote_char)
            writer.writerows(lines)


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

    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active

    res_keys = [[ind] + [i.value for k, i in enumerate(row, 1) if k in key_cols]
                for ind, row in enumerate(ws.rows, 1)
                if ind >= begin_row and row[sign_col - 1].value == sign]

    wb.close()
    return res_keys


def _get_csv_keys(path: str,
                  begin_row: int,
                  sign_col: Union[int, str],
                  sign: str,
                  key_cols: Union[list, tuple],
                  encoding: str,
                  delimiter: str,
                  quotechar: str) -> List[list]:
    """返回key列内容，第一位为行号，其余为key列的值       \n
    eg.[3, '名称', 'id']
    :param path: 文件路径
    :param begin_row: 数据起始行
    :param sign_col: 用于判断是否已填数据的列，从1开始
    :param sign: 按这个值判断是否已填数据
    :param key_cols: 关键字所在列，可以是多列
    :param encoding: 字符编码
    :param delimiter: 分隔符
    :param quotechar: 引用符
    :return: 关键字组成的列表
    """
    key_cols = [i - 1 if isinstance(i, int) else column_index_from_string(i) - 1 for i in key_cols]
    sign_col = column_index_from_string(sign_col) if isinstance(sign_col, str) else sign_col
    sign = '' if sign is None else str(sign)
    sign_col -= 1
    begin_row -= 1
    res_keys = []

    with open(path, 'r', encoding=encoding) as f:
        reader = csv_reader(f, delimiter=delimiter, quotechar=quotechar)
        lines = list(reader)[begin_row:]

        for k, line in enumerate(lines):
            row_sign = '' if sign_col > len(line) - 1 else line[sign_col]
            if row_sign == sign:
                key = [k + 1]
                res_keys.append(key + [i for num, i in enumerate(line) if num in key_cols])

    return res_keys
