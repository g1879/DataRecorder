# -*- coding:utf-8 -*-

from pathlib import Path
from typing import Union


def align_csv(path: Union[str, Path],
              encoding: str = 'utf-8',
              delimiter: str = ',',
              quotechar: str = '"') -> None: ...


def get_usable_path(path: Union[str, Path]) -> Path: ...


def make_valid_file_name(full_name: str) -> str: ...
