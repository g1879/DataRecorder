# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, List, Any, Tuple

from openpyxl.styles import Font

from .base import BaseRecorder


class Filler(BaseRecorder):

    def __init__(self,
                 path: Union[str, Path],
                 cache_size: int = None,
                 key_cols: Union[str, int, list, tuple, bool] = True,
                 begin_row: int = 2,
                 sign_col: Union[str, int, bool] = True,
                 data_col: Union[int, str] = None,
                 sign: str = None,
                 deny_sign: bool = False) -> None:
        self._key_cols: Union[List[int], bool] = ...
        self._begin_row: Union[str, int] = ...
        self._sign_col: Union[int, None, bool] = ...
        self._data_col: int = ...
        self._sign: str = ...
        self._deny_sign: bool = ...
        self._link_font: Font = ...

    @property
    def sign(self) -> str: ...

    def set_sign(self, value) -> None: ...

    @property
    def deny_sign(self) -> bool: ...

    def set_deny_sign(self, on_off=True) -> bool: ...

    @property
    def key_cols(self) -> Union[List[int], bool]: ...

    def set_key_cols(self, cols: Union[str, int, list, tuple, bool]) -> None: ...

    @property
    def sign_col(self) -> Union[int, None, bool]: ...

    def set_sign_col(self, col: Union[str, int, bool]) -> None: ...

    @property
    def data_col(self) -> int: ...

    def set_data_col(self, col: Union[str, int]) -> None: ...

    @property
    def begin_row(self) -> Union[str, int]: ...

    def set_begin_row(self, row: int) -> None: ...

    @property
    def keys(self) -> list: ...

    def set_path(self,
                 path: Union[str, Path],
                 key_cols: Union[str, int, list, tuple] = None,
                 begin_row: int = None,
                 sign_col: Union[str, int] = None,
                 data_col: int = None,
                 sign: Union[int, float, str] = None,
                 deny_sign: bool = None) -> None: ...

    def add_data(self, data: Any,
                 coord: Union[list, Tuple[Union[None, int, str], Union[int, str]], str, int] = 'newline') -> None: ...

    def set_link(self,
                 coord: Union[int, str, tuple, list],
                 link: str,
                 content: Union[int, str, float] = None) -> None: ...

    def set_link_style(self, style: Font) -> None: ...

    def _record(self) -> None: ...

    def _to_xlsx(self) -> None: ...

    def _to_csv(self) -> None: ...

# def _get_xlsx_keys(path: str,
#                    begin_row: int,
#                    sign_col: Union[int, str, None],
#                    sign: Union[int, float, str],
#                    key_cols: Union[list, tuple],
#                    deny_sign: bool,
#                    table: str) -> List[list]:...
#
#
# def _get_csv_keys(path: str,
#                   begin_row: int,
#                   sign_col: Union[int, str, None],
#                   sign: Union[int, float, str],
#                   key_cols: Union[list, tuple],
#                   encoding: str,
#                   delimiter: str,
#                   quotechar: str,
#                   deny_sign: bool) -> List[list]:...
