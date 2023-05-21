# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, Any

from .recorder import Recorder
from .setter import DBSetter


class DBRecorder(Recorder):
    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = None,
                 table: str = None):
        self._conn = ...
        self._cur = ...

    @property
    def set(self) -> DBSetter: ...

    def __del__(self): ...

    def add_data(self, data: Any, table: str = None) -> None: ...

    def _connect(self) -> None: ...

    def _close_connection(self) -> None: ...

    def _record(self) -> None: ...

    def _to_sqlite(self) -> None: ...
