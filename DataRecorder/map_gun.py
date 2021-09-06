# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union

from .base import BaseRecorder


class MapGun(BaseRecorder):
    SUPPORTS = ('xlsx', 'csv')

    def __init__(self, path: Union[str, Path]):
        super().__init__(path)

    def add_data(self, data):
        pass

    def _record(self):
        pass
