# -*- coding:utf-8 -*-
from typing import Union, Any

from .base import BaseRecorder
from .cell_style import CellStyle
from .setter import RecorderSetter


class Recorder(BaseRecorder):
    _row_styles: Union[list, None] = ...
    _col_height: Union[float, None] = ...
    _follow_styles: bool = ...
    _row_styles_len: Union[int, None] = ...
    _style: Union[CellStyle, None] = ...

    @property
    def set(self) -> RecorderSetter: ...

    def add_data(self, data: Any) -> None: ...

    def _record(self) -> None: ...

    def _to_xlsx(self) -> None: ...

    def _to_csv(self) -> None: ...

    def _to_txt(self) -> None: ...

    def _to_json(self) -> None: ...

    def _data_to_list_or_dict(self, data: Union[list, tuple, dict]) -> Union[list, dict]: ...