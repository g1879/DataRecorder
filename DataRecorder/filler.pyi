# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, List, Any, Tuple, Optional

from .base import BaseRecorder
from .style import CellStyle
from .setter import FillerSetter


class Filler(BaseRecorder):
    _set: FillerSetter = ...
    _key_cols: Optional[List[int], bool] = ...
    _begin_row: Optional[str, int] = ...
    _sign_col: Optional[int, bool] = ...
    _data_col: Optional[int] = ...
    _sign: Any = ...
    _deny_sign: bool = ...
    _link_style: CellStyle = ...
    _quote_char: str = ...
    _delimiter: str = ...
    _data: Union[list, dict] = ...
    data: Union[list, dict] = ...
    row_num_title: str = ...

    def __init__(self, path: Optional[str, Path] = None,
                 cache_size: int = None,
                 key_cols: Union[str, int, list, tuple, bool] = True,
                 begin_row: int = 2,
                 sign_col: Union[str, int, bool] = True,
                 data_col: Optional[int, str] = None,
                 sign: Any = None,
                 deny_sign: bool = False) -> None: ...

    @property
    def sign(self) -> str: ...

    @property
    def deny_sign(self) -> bool: ...

    @property
    def key_cols(self) -> Union[List[int], bool]: ...

    @property
    def sign_col(self) -> Optional[int, bool]: ...

    @property
    def data_col(self) -> int: ...

    @property
    def begin_row(self) -> Union[str, int]: ...

    @property
    def keys(self) -> list: ...

    @property
    def dict_keys(self) -> List[dict]: ...

    @property
    def set(self) -> FillerSetter: ...

    @property
    def delimiter(self) -> str: ...

    @property
    def quote_char(self) -> str: ...

    def add_data(self, data: Any,
                 coord: Union[list, Tuple[Optional[int, str], Union[int, str]], str, int] = 'newline',
                 table: Union[str, bool] = None) -> None: ...

    def set_link(self,
                 coord: Union[int, str, tuple, list],
                 link: Optional[str],
                 content: Optional[int, str, float] = None) -> None: ...

    def set_style(self, coord: Union[int, str, tuple, list], style: Optional[CellStyle],
                  replace: bool = True) -> None: ...

    def set_img(self, coord: Union[int, str, tuple, list], img_path: Optional[str, Path], width: float = None,
                height: float = None) -> None: ...

    def set_row_height(self, row: int, height: float) -> None: ...

    def set_col_width(self, col: Union[int, str], width: float) -> None: ...

    def _record(self) -> None: ...

    def _to_xlsx(self) -> None: ...

    def _to_csv(self) -> None: ...


def get_xlsx_keys(filler: Filler, as_dict: bool) -> List[Union[list, dict]]: ...


def get_csv_keys(filler: Filler, as_dict: bool) -> List[Union[list, dict]]: ...
