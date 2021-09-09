# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union

from .base import BaseRecorder


class MapGun(BaseRecorder):
    """把二维数据填充到以左上角坐标为起点的范围"""

    def __init__(self, path: Union[str, Path]):
        super().__init__(path)

    def add_data(self, data):
        """接收二维数据，若是一维的，每个元素作为一行看待"""
        pass

    def _record(self):
        pass
