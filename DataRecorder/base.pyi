# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from threading import Lock
from typing import Union, Any, Optional

from .setter import OriginalSetter, BaseSetter


class OriginalRecorder(object):
    SUPPORTS: tuple = ...
    _cache: int = ...
    _path: str = ...
    _type: str = ...
    _data: Union[list, dict] = ...
    _lock: Lock = ...
    _pause_add: bool = ...
    _pause_write: bool = ...
    show_msg: bool = ...
    _setter: OriginalSetter = ...
    _data_count: int = ...

    def __init__(self,
                 path: Optional[str, Path] = None,
                 cache_size: int = None) -> None: ...

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
    def data(self) -> Union[dict, list]: ...

    def record(self, new_path: Optional[str, Path] = None) -> str: ...

    def clear(self) -> None: ...

    @abstractmethod
    def add_data(self, data): ...

    @abstractmethod
    def _record(self): ...


class BaseRecorder(OriginalRecorder):
    SUPPORTS : tuple = ...
    _encoding: str = ...
    _before: list = ...
    _after: list = ...
    _table: Union[str, bool] = ...
    _setter: BaseSetter = ...

    def __init__(self, path: Optional[str, Path] = None, cache_size: int = None) -> None: ...

    @property
    def set(self) -> BaseSetter: ...

    @property
    def before(self) -> Any: ...

    @property
    def after(self) -> Any: ...

    @property
    def table(self) -> Union[str, bool]: ...

    @property
    def encoding(self) -> str: ...

    @abstractmethod
    def _record(self): ...

    def _data_to_list(self, data: Union[list, tuple, dict], long: int = None) -> list: ...
