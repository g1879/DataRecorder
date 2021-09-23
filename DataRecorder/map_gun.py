# -*- coding:utf-8 -*-
from csv import reader as csv_reader, writer as csv_writer
from pathlib import Path
from typing import Union

from openpyxl import load_workbook, Workbook

from .base import BaseRecorder, _data_to_list, _parse_coord, _process_content


class MapGun(BaseRecorder):
    """把二维数据填充到以左上角坐标为起点的范围"""

    def __init__(self, path: Union[str, Path],
                 coord: Union[str, tuple, list] = None,
                 float_coord: bool = True):
        """初始化                                              \n
        :param path: 保存的文件路径
        :param coord: 目标左上角坐标
        :param float_coord: 保存数据后坐标是否移动到底部
        """
        super().__init__(path, 1)
        self.coord = coord or [1, 1]
        self.float_coord = float_coord

    @property
    def coord(self) -> list:
        """起始坐标"""
        return self._coord

    @coord.setter
    def coord(self, coord: Union[str, tuple, list]) -> None:
        """设置填写坐标                                                               \n
        :param coord: 接受几种形式：'A3', '3,1', (3, 1), [3, 1]，除第一种外都是行在前
        :return: None
        """
        self._coord = list(_parse_coord(coord, None, int))

    @property
    def cache_size(self) -> int:
        """返回缓存大小"""
        return self._cache

    @cache_size.setter
    def cache_size(self, cache_size: int) -> None:
        """固定缓存大小                   \n
        :param cache_size: 缓存大小
        :return: None
        """
        pass

    def add_data(self, data: Union[list, tuple], coord: Union[str, tuple, list] = None) -> None:
        """接收二维数据，若是一维的，每个元素作为一行看待    \n
        :param data: 二维数据
        :param coord: 左上角坐标
        :return: None
        """
        if coord is not None:
            self.coord = coord
        self._data = data
        self.record()

    def _record(self) -> None:
        """记录数据"""
        if self.type == 'xlsx':
            _record_to_xlsx(self.path, self._data, self.coord, self._before, self._after)
        elif self.type == 'csv':
            _record_to_csv(self.path, self._data, self.coord, self._before, self._after, self.encoding,
                           self.delimiter, self.quote_char)

        if self.float_coord:
            self.coord[0] += len(self._data)


def _record_to_xlsx(file_path: str,
                    data: list,
                    coord: list,
                    before: Union[list, tuple, dict] = None,
                    after: Union[list, tuple, dict] = None) -> None:
    """记录数据到xlsx文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param coord: 左上角坐标
    :param before: 数据前面的列
    :param after: 数据后面的列
    :return: None
    """
    wb = load_workbook(file_path) if Path(file_path).exists() else Workbook()
    ws = wb.active

    row, col = coord
    for i in data:
        if not isinstance(i, (list, tuple)):
            i = (i,)
        now_data = _data_to_list(i, before, after)

        for ind, j in enumerate(now_data):
            ws.cell(row, col + ind).value = _process_content(j)

        row += 1

    wb.save(file_path)
    wb.close()


def _record_to_csv(file_path: str,
                   data: Union[list, tuple],
                   coord: list,
                   before: Union[list, dict] = None,
                   after: Union[list, dict] = None,
                   encoding: str = 'utf-8',
                   delimiter: str = ',',
                   quotechar: str = '"') -> None:
    """填写数据到xlsx文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param coord: 左上角坐标
    :param before: 数据前面的列
    :param after: 数据后面的列
    :param encoding: 字符编码
    :param delimiter: 分隔符
    :param quotechar: 引用符
    :return: None
    """
    if not Path(file_path).exists():
        with open(file_path, 'w', encoding=encoding):
            pass

    with open(file_path, 'r', encoding=encoding) as f:
        reader = csv_reader(f, delimiter=delimiter, quotechar=quotechar)
        lines = list(reader)
        lines_len = len(lines)
        row, col = coord

        # 若行数不够，填充行数
        for _ in range(row + len(data) - lines_len):
            lines.append([])

        # 填入数据
        for ind, i in enumerate(data):
            if not isinstance(i, (list, tuple)):
                i = (i,)

            now_data = _data_to_list(i, before, after)
            row_num = row + ind - 1

            # 若列数不够，填充空列
            lines[row_num].extend([None] * (col - len(lines[row_num]) + len(now_data) - 1))

            # 填充一行数据
            for k, j in enumerate(now_data):
                lines[row_num][col + k - 1] = _process_content(j)

        writer = csv_writer(open(file_path, 'w', encoding=encoding, newline=''),
                            delimiter=delimiter, quotechar=quotechar)
        writer.writerows(lines)
