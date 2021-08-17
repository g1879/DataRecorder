# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union

from functions import _data_to_list
from .base import BaseRecorder


class Filler(BaseRecorder):
    SUPPORTS = ('xlsx',)

    def __init__(self, file_path: Union[str, Path],
                 cache_size: int = 50,
                 key_col: Union[str, int] = None,
                 begin_row: int = 2,
                 sign_col: Union[str, int] = None
                 ):
        """初始化                                  \n
        :param file_path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件
        """
        super().__init__(file_path, cache_size)
        self.key_col = key_col

    @property
    def key_col(self):
        return self._key_col

    @key_col.setter
    def key_col(self, col: Union[str, int]):
        self._key_col = col

    # @property
    # def keys(self):
    #     if self.file_type == 'xlsx':
    #

    def set_file_path(self, path: Union[str, Path], key_col: Union[str, int]):
        if not Path(path).exists():
            raise FileNotFoundError('文件不存在')
        self.file_path = path
        self.key_col = key_col

    def add_data(self, data):
        pass

    def record(self):
        pass


def _fill_to_xlsx(file_path: str,
                  data: list,
                  before: Union[list, dict] = None,
                  after: Union[list, dict] = None,
                  col: int = None,
                  row: int = None) -> None:
    """记录数据到xlsx文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :return: None
    """
    try:
        import openpyxl

    except ModuleNotFoundError:
        import os
        os.system('pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl')
        import openpyxl

    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    for i in data:
        for key, j in enumerate(_data_to_list(i[1], before, after)):
            ws.cell(i[0], col + key).value = j

    wb.save(file_path)
    wb.close()
