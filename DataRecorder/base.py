# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from typing import Union


class BaseRecorder(object):
    """记录器的父类"""
    SUPPORTS = ()

    def __init__(self, path: Union[str, Path], cache_size: int = 50):
        """初始化                                  \n
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件
        """
        self._data = []
        self._before = []
        self._after = []
        self._type = None
        self.path = path
        self.cache_size = cache_size
        self.encoding: str = 'utf-8'

    def __del__(self) -> None:
        """对象关闭时把剩下的数据写入文件"""
        if self._type == 'xlsx' and self._data:
            print(f'{self._data}\n\n以上数据未保存')
        else:
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
        if not isinstance(cache_size, int) or cache_size < 0:
            raise TypeError('cache_size值只能是int，且必须>=0')
        self._cache = cache_size

    @property
    def path(self) -> str:
        """返回文件路径"""
        return self._path

    @path.setter
    def path(self, path: Union[str, Path]) -> None:
        """设置文件路径                \n
        :param path: 文件路径
        :return: None
        """
        if isinstance(path, str):
            self._type = path.split('.')[-1]
        elif isinstance(path, Path):
            self._type = path.suffix[1:]
        else:
            raise TypeError(f'参数file_path只能是str或Path，非{type(path)}。')

        if self._type not in self.SUPPORTS:
            raise TypeError(f'只支持{"、".join(self.SUPPORTS)}格式文件。')

        self.record()  # 更换文件前自动记录剩余数据
        self._path = str(path) if isinstance(path, Path) else path

    @property
    def type(self) -> str:
        return self._type

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


def _data_to_list(data: Union[list, tuple, dict],
                  before: Union[list, dict, None] = None,
                  after: Union[list, dict, None] = None) -> list:
    """将传入的数据转换为列表形式          \n
    :param data: 要处理的数据
    :param before: 数据前的列
    :param after: 数据后的列
    :return: 转变成列表方式的数据
    """
    return_list = []
    if data is not None and not isinstance(data, (list, tuple, dict)):
        data = [data]
    if before is not None and not isinstance(before, (list, tuple, dict)):
        before = [before]
    if after is not None and not isinstance(after, (list, tuple, dict)):
        after = [after]

    for i in (before, data, after):
        if isinstance(i, dict):
            return_list.extend(list(i.values()))
        elif i is None:
            pass
        elif isinstance(i, list):
            return_list.extend(i)
        elif isinstance(i, tuple):
            return_list.extend(list(i))
        else:
            return_list.extend([str(i)])

    return return_list


def _data_to_list_or_dict(data: Union[list, tuple, dict],
                          before: Union[list, tuple, dict, None] = None,
                          after: Union[list, tuple, dict, None] = None) -> Union[list, dict]:
    """将传入的数据转换为列表或字典形式，用于记录到txt或json          \n
    :param data: 要处理的数据
    :param before: 数据前的列
    :param after: 数据后的列
    :return: 转变成列表或字典形式的数据
    """
    if isinstance(data, (list, tuple)):
        return _data_to_list(data, before, after)

    elif isinstance(data, dict):
        if isinstance(before, dict):
            data = {**before, **data}

        if isinstance(after, dict):
            data = {**data, **after}

        return data
