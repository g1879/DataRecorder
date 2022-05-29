# -*- coding:utf-8 -*-
from pathlib import Path
from time import sleep
from typing import Union, List, Tuple

from .base import OriginalRecorder


class ByteRecorder(OriginalRecorder):
    """用于记录字节数据的工具"""
    SUPPORTS = ('any',)

    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = None):
        super().__init__(path, cache_size)

    def add_data(self,
                 data: Union[bytes, List[bytes], Tuple[bytes, ...]]) -> None:
        """添加数据，可添加多个                   \n
        :param data: bytes或bytes组成的列表
        :return: None
        """
        while self._pause_add:
            sleep(.1)

        if isinstance(data, bytes):
            self._data.append(data)

        elif isinstance(data, (list, tuple)):
            if any([i for i in data if not isinstance(i, bytes)]):
                raise TypeError('只能接受bytes类型数据。')
            if isinstance(data, tuple):
                data = list(data)

            self._data.extend(data)

        else:
            raise TypeError('只能接受bytes类型数据。')

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def _record(self) -> None:
        """记录数据到文件"""
        with open(self.path, 'ab+') as f:
            f.write(b''.join(self._data))
