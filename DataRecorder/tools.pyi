# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, Tuple, Any, Optional

from .base import BaseRecorder


def align_csv(path: Union[str, Path], encoding: str = 'utf-8', delimiter: str = ',', quotechar: str = '"') -> None: ...


def get_usable_path(path: Union[str, Path]) -> Path: ...


def make_valid_file_name(full_name: str) -> str: ...


def parse_coord(coord: Optional[int, str, list, tuple] = None, data_col: int = None) -> Tuple[Optional[int], int]: ...


def process_content(content: Any, excel: bool = False) -> Optional[int, str, float]: ...


def ok_list(data_list: Union[list, dict], excel: bool = False, as_str: bool = False) -> list: ...


def get_usable_coord(coord: Union[tuple, list], max_row: int, max_col: int) -> Tuple[int, int]: ...


def data_to_list_or_dict(recorder: BaseRecorder, data: Union[list, tuple, dict]) -> Union[list, dict]: ...
