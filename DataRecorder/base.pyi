# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from threading import Lock
from typing import Union, Any, Tuple


class OriginalRecorder(object):
    SUPPORTS: tuple = ...

    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = None) -> None:
        self._cache: int = ...
        self._path: str = ...
        self._type: str = ...
        self._data: list = ...
        self._lock: Lock = ...
        self._pause_add: bool = ...
        self._pause_write: bool = ...
        self.show_msg: bool = ...

    def __del__(self) -> None: ...

    @property
    def cache_size(self) -> int: ...

    def set_cache_size(self, cache_size: int) -> None: ...

    @property
    def path(self) -> str: ...

    def set_path(self, path: Union[str, Path], file_type: str = None) -> None: ...

    @property
    def type(self) -> str: ...

    def set_type(self, file_type: str) -> None: ...

    @property
    def data(self) -> list: ...

    def record(self, new_path: Union[str, Path] = None) -> Union[str, list]: ...

    def clear(self) -> None: ...

    @abstractmethod
    def add_data(self, data): ...

    @abstractmethod
    def _record(self): ...


class BaseRecorder(OriginalRecorder):
    SUPPORTS = ('xlsx', 'csv')

    def __init__(self, path: Union[str, Path] = None, cache_size: int = None) -> None:
        self._encoding: str = ...
        self._delimiter: str = ...
        self._quote_char: str = ...
        self._before: list = ...
        self._after: list = ...
        self._table: str = ...

    @property
    def before(self) -> Any: ...

    @property
    def after(self) -> Any: ...

    @property
    def table(self) -> str: ...

    def set_table(self, table: str) -> None: ...

    @property
    def encoding(self) -> str: ...

    @property
    def delimiter(self) -> str: ...

    @property
    def quote_char(self) -> str: ...

    def set_encoding(self, encoding) -> None: ...

    def set_delimiter(self, delimiter) -> None: ...

    def set_quote_char(self, quote_char) -> None: ...

    def set_before(self, before: Any) -> None: ...

    def set_after(self, after: Any) -> None: ...

    def set_head(self, head: Union[list, tuple]) -> None: ...

    @abstractmethod
    def add_data(self, data): ...

    @abstractmethod
    def _record(self): ...

    def _data_to_list(self, data: Union[list, tuple, dict]) -> list: ...


def parse_coord(coord: Union[int, str, list, tuple, None] = None,
                data_col: int = None) -> Tuple[Union[int, None], int]: ...


def process_content(content: Any, excel: bool = False) -> Union[int, str, float, None]: ...


def ok_list(data_list: Union[list, dict], excel: bool = False, as_str: bool = False) -> list: ...


def get_usable_coord(coord: Union[tuple, list], max_row: int, max_col: int) -> Tuple[int, int]: ...
