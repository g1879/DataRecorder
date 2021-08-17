# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from typing import Union


class BaseRecorder(object):
    SUPPORTS = ()

    def __init__(self, file_path: Union[str, Path], cache_size: int = 50):
        """初始化                                  \n
        :param file_path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件
        """
        self._data = []
        self._before = []
        self._after = []
        self._file_type = None
        self.file_path = file_path
        self.cache_size = cache_size
        self.encoding: str = 'utf-8'

    def __del__(self) -> None:
        """对象关闭时把剩下的数据写入文件"""
        if not self.file_path.endswith('.xlsx'):
            self.record()

    @property
    def cache_size(self) -> int:
        """返回缓存大小"""
        return self._cache

    @cache_size.setter
    def cache_size(self, cache_size: int) -> None:
        """设置缓存大小                   \n
        :param cache_size: 缓存大小
        :return: None
        """
        if not isinstance(cache_size, int):
            raise TypeError('参数cache_size只能为int。')
        self._cache = cache_size

    @property
    def file_path(self) -> str:
        """返回文件路径"""
        return self._file_path

    @file_path.setter
    def file_path(self, path: Path[str, Path]) -> None:
        """设置文件路径                \n
        :param path: 文件路径
        :return: None
        """
        if isinstance(path, str):
            self._file_type = path.split('.')[-1]
        elif isinstance(path, Path):
            self._file_type = path.suffix[1:]
        else:
            raise TypeError(f'参数file_path只能是str或Path，非{type(path)}。')

        if self._file_type not in self.SUPPORTS:
            raise TypeError(f'只支持{"、".join(self.SUPPORTS)}格式文件。')

        self.record()  # 更换文件前自动记录剩余数据
        self._file_path = str(path)

    @property
    def file_type(self) -> str:
        return self._file_type

    def set_before(self, before: Union[list, tuple, str, dict]) -> None:
        """设置在数据前面补充的列                              \n
        :param before: 列表、元组或字符串，为字符串时则补充一列
        :return: None
        """
        self.record()

        if before is None:
            self._before = []
        elif isinstance(before, (list, dict)):
            self._before = before
        elif isinstance(before, tuple):
            self._before = list(before)
        else:
            self._before = [str(before)]

    def set_after(self, after: Union[list, tuple, str, dict]) -> None:
        """设置在数据后面补充的列                                \n
        :param after: 列表、元组或字符串，为字符串时则补充一列
        :return: None
        """
        self.record()

        if after is None:
            self._after = []
        elif isinstance(after, (list, dict)):
            self._after = after
        elif isinstance(after, tuple):
            self._after = list(after)
        else:
            self._before = [str(after)]

    def clear(self) -> None:
        """清空缓存中的数据"""
        self._data = []

    @abstractmethod
    def record(self):
        pass

    @abstractmethod
    def add_data(self, data):
        pass
