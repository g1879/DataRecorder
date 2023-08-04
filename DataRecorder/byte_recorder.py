# -*- coding:utf-8 -*-
from pathlib import Path
from time import sleep

from .base import OriginalRecorder


class ByteRecorder(OriginalRecorder):
    SUPPORTS = ('any',)
    __END = (0, 2)

    def __init__(self, path=None, cache_size=None):
        """用于记录字节数据的工具
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        super().__init__(path, cache_size)

    def add_data(self, data, seek=None):
        """添加一段二进制数据
        :param data: bytes类型数据
        :param seek: 在文件中的位置，None表示最后
        :return: None
        """
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.2)

        if not isinstance(data, bytes):
            raise TypeError('只能接受bytes类型数据。')
        if seek is not None and not (isinstance(seek, int) and seek >= 0):
            raise ValueError('seek参数只能接受None或大于等于0的整数。')

        self._data.append((data, seek))

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def _record(self):
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
