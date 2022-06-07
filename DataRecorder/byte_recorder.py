# -*- coding:utf-8 -*-
from pathlib import Path
from time import sleep
from typing import Union

from .base import OriginalRecorder


class ByteRecorder(OriginalRecorder):
    """用于记录字节数据的工具"""
    SUPPORTS = ('any',)
    __END = (0, 2)

    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = None):
        super().__init__(path, cache_size)

    def add_data(self,
                 data: bytes,
                 seek: int = None) -> None:
        """添加一段二进制数据                      \n
        :param data: bytes或bytes组成的列表
        :param seek: 在文件中的位置，None表示最后
        :return: None
        """
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.1)

        if not isinstance(data, bytes):
            raise TypeError('只能接受bytes类型数据。')
        if seek is not None and not (isinstance(seek, int) and seek >= 0):
            raise ValueError('seek参数只能接受None或大于等于0的整数。')

        self._data.append((data, seek))

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def _record(self) -> None:
        """记录数据到文件"""
        if not Path(self.path).exists():
            with open(self.path, 'w'):
                pass

        with open(self.path, 'rb+') as f:
            previous = None
            for i in self._data:
                loc = ByteRecorder.__END if i[1] is None else (i[1], 0)
                if not (previous == loc == ByteRecorder.__END):
                    f.seek(loc[0], loc[1])
                    previous = loc
                f.write(i[0])
