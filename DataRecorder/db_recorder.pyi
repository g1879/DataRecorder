# -*- coding:utf-8 -*-
from pathlib import Path
from sqlite3 import Connection, Cursor
from typing import Union, Any, Optional

from .base import BaseRecorder
from .setter import DBSetter


class DBRecorder(BaseRecorder):
    _conn: Optional[Connection] = ...
    _cur: Optional[Cursor] = ...
    _setter: Optional[DBSetter] = ...
    _data: dict = ...
    data: dict = ...

    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = None,
                 table: str = None): ...

    @property
    def set(self) -> DBSetter: ...

    def __del__(self): ...

    def add_data(self, data: Any, table: str = None) -> None: ...

    def run_sql(self, sql: str, single: bool = True, commit: bool = False) -> Optional[list, tuple]: ...

    def _connect(self) -> None: ...

    def _close_connection(self) -> None: ...

    def _record(self) -> None: ...

    def _to_database(self, data_list: list, table: str, tables: dict) -> None: ...
