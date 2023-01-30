# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union

from .base import OriginalRecorder


class ByteRecorder(OriginalRecorder):
    SUPPORTS: tuple = ...
    __END: tuple = ...

    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = None): ...

    def add_data(self,
                 data: bytes,
                 seek: int = None) -> None: ...

    def _record(self) -> None: ...
