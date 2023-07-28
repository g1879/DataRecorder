# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from threading import Lock
from typing import Union, Any, Optional

from .setter import OriginalSetter, BaseSetter


class OriginalRecorder(object):
    SUPPORTS: tuple = ...

    def __init__(self,
                 path: Optional[str, Path] = None,
                 cache_size: int = None) -> None:
        self._cache: int = ...
        self._path: str = ...
        self._type: str = ...
        self._data: list = ...
        self._lock: Lock = ...
        self._pause_add: bool = ...
        self._pause_write: bool = ...
        self.show_msg: bool = ...
        self._setter: OriginalSetter = ...

    def __del__(self) -> None: ...

    @property
    def set(self) -> OriginalSetter: ...

    @property
    def cache_size(self) -> int: ...

    @property
    def path(self) -> str: ...

    @property
    def type(self) -> str: ...

    @property
    def data(self) -> list: ...

    def record(self, new_path: Optional[str, Path] = None) -> Union[str, list]: ...

    def clear(self) -> None: ...

    @abstractmethod
    def add_data(self, data): ...

    @abstractmethod
    def _record(self): ...


class BaseRecorder(OriginalRecorder):
    SUPPORTS = ('xlsx', 'csv')

    def __init__(self, path: Optional[str, Path] = None, cache_size: int = None) -> None:
        self._encoding: str = ...
        self._before: list = ...
        self._after: list = ...
        self._table: str = ...
        self._setter: BaseSetter = ...

    @property
    def set(self) -> BaseSetter: ...

    @property
    def before(self) -> Any: ...

    @property
    def after(self) -> Any: ...

    @property
    def table(self) -> str: ...

    @property
    def encoding(self) -> str: ...

    @abstractmethod
    def _record(self): ...

    def _data_to_list(self, data: Union[list, tuple, dict]) -> list: ...
