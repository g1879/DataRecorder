# -*- coding:utf-8 -*-
from threading import Lock
from typing import Literal, Optional, Any, Union

from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Font, Border, Fill, Protection, Side, PatternFill

LINES = Literal['dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'medium', 'mediumDashDot',
'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin', None]


class CellStyle(object):
    font_args: tuple = ...
    border_args: tuple = ...
    alignment_args: tuple = ...
    protection_args: tuple = ...
    gradient_fill_args: tuple = ...
    pattern_fill_args: tuple = ...

    def __init__(self) -> None:
        self._font: Optional[CellFont] = ...
        self._border: Optional[CellBorder] = ...
        self._pattern_fill: Optional[CellPatternFill] = ...
        self._gradient_fill: Optional[CellGradientFill] = ...
        self._number_format: Optional[CellNumberFormat] = ...
        self._protection: Optional[CellProtection] = ...
        self._alignment: Optional[CellAlignment] = ...
        self._Font: Optional[Font] = None
        self._Border: Optional[Border] = None
        self._Alignment: Optional[Alignment] = None
        self._Fill: Optional[Fill] = None
        self._Protection: Optional[Protection] = None

    @property
    def font(self) -> CellFont: ...

    @property
    def border(self) -> CellBorder: ...

    @property
    def alignment(self) -> CellAlignment: ...

    @property
    def pattern_fill(self) -> CellPatternFill: ...

    @property
    def gradient_fill(self) -> CellGradientFill: ...

    @property
    def number_format(self) -> CellNumberFormat: ...

    @property
    def protection(self) -> CellProtection: ...

    def to_cell(self, cell: Cell, replace: bool = True) -> None: ...

    def _cover_to_cell(self, cell: Cell) -> None: ...

    def _replace_to_cell(self, cell: Cell) -> None: ...


def _handle_args(args: tuple, src: Any, target: Any) -> dict: ...


class CellFont(object):
    _LINE_STYLES: tuple = ...
    _SCHEMES: tuple = ...
    _VERT_ALIGNS: tuple = ...

    def __init__(self):
        self.name: str = ...
        self.charset: int = ...
        self.size: float = ...
        self.bold: bool = ...
        self.italic: bool = ...
        self.strike: bool = ...
        self.outline: bool = ...
        self.shadow: bool = ...
        self.condense: bool = ...
        self.extend: bool = ...
        self.underline: Literal['single', 'double', 'singleAccounting', 'doubleAccounting'] = ...
        self.vertAlign: Literal['superscript', 'subscript', 'baseline'] = ...
        self.color: str = ...
        self.scheme: Literal['major', 'minor'] = ...

    def set_name(self, name: Optional[str]) -> None: ...

    def set_charset(self, charset: Optional[int]) -> None: ...

    def set_size(self, size: Optional[float]) -> None: ...

    def set_bold(self, on_off: Optional[bool]) -> None: ...

    def set_italic(self, on_off: Optional[bool]) -> None: ...

    def set_strike(self, on_off: Optional[bool]) -> None: ...

    def set_outline(self, on_off: Optional[bool]) -> None: ...

    def set_shadow(self, on_off: Optional[bool]) -> None: ...

    def set_condense(self, on_off: Optional[bool]) -> None: ...

    def set_extend(self, on_off: Optional[bool]) -> None: ...

    def set_color(self, color: Optional[str, tuple]) -> None: ...

    def set_underline(self,
                      option: Literal['single', 'double', 'singleAccounting', 'doubleAccounting', None]) -> None: ...

    def set_vertAlign(self, option: Literal['superscript', 'subscript', 'baseline', None]) -> None: ...

    def set_scheme(self, option: Literal['major', 'minor', None]) -> None: ...


class CellBorder(object):
    _LINE_STYLES: tuple = ...

    def __init__(self):
        self.start: Side = ...
        self.end: Side = ...
        self.left: Side = ...
        self.right: Side = ...
        self.top: Side = ...
        self.bottom: Side = ...
        self.diagonal: Side = ...
        self.vertical: Side = ...
        self.horizontal: Side = ...
        self.horizontal: Side = ...
        self.outline: bool = ...
        self.diagonalUp: bool = ...
        self.diagonalDown: bool = ...

    def set_start(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_end(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_left(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_right(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_top(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_bottom(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_diagonal(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_vertical(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_horizontal(self, style: LINES, color: Optional[str, tuple]) -> None: ...

    def set_outline(self, on_off: bool) -> None: ...

    def set_diagonalDown(self, on_off: bool) -> None: ...

    def set_diagonalUp(self, on_off: bool) -> None: ...


class CellAlignment(object):
    _horizontal_alignments: tuple = ...
    _vertical_alignments: tuple = ...

    def __init__(self):
        self.horizontal = 'notSet'
        self.vertical = 'notSet'
        self.indent = 'notSet'
        self.relativeIndent = 'notSet'
        self.justifyLastLine = 'notSet'
        self.readingOrder = 'notSet'
        self.textRotation = 'notSet'
        self.wrapText = 'notSet'
        self.shrinkToFit = 'notSet'

    def set_horizontal(self,
                       horizontal: Literal['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous',
                       'distributed', None]) -> None: ...

    def set_vertical(self, vertical: Literal['top', 'center', 'bottom', 'justify', 'distributed', None]) -> None: ...

    def set_indent(self, indent: int) -> None: ...

    def set_relativeIndent(self, indent: int) -> None: ...

    def set_justifyLastLine(self, on_off: Optional[bool]) -> None: ...

    def set_readingOrder(self, value: int) -> None: ...

    def set_textRotation(self, value: int) -> None: ...

    def set_wrapText(self, on_off: Optional[bool]) -> None: ...

    def set_shrinkToFit(self, on_off: Optional[bool]) -> None: ...


class CellGradientFill(object):
    def __init__(self):
        self.type: str = ...
        self.degree: float = ...
        self.left: float = ...
        self.right: float = ...
        self.top: float = ...
        self.bottom: float = ...
        self.stop: Union[list, tuple] = ...

    def set_type(self, name: Literal['linear', 'path']) -> None: ...

    def set_degree(self, value: float) -> None: ...

    def set_left(self, value: float) -> None: ...

    def set_right(self, value: float) -> None: ...

    def set_top(self, value: float) -> None: ...

    def set_bottom(self, value: float) -> None: ...

    def set_stop(self, values: Union[list, tuple]) -> None: ...


class CellPatternFill(object):
    _FILES: tuple = ...

    def __init__(self):
        self.patternType: str = ...
        self.fgColor: str = ...
        self.bgColor: str = ...

    def set_patternType(self, name: Literal[
        'none', 'solid', 'darkDown', 'darkGray', 'darkGrid', 'darkHorizontal', 'darkTrellis', 'darkUp',
        'darkVertical', 'gray0625', 'gray125', 'lightDown', 'lightGray', 'lightGrid', 'lightHorizontal',
        'lightTrellis', 'lightUp', 'lightVertical', 'mediumGray', None]) -> None: ...

    def set_fgColor(self, color: Optional[str, tuple]) -> None: ...

    def set_bgColor(self, color: Optional[str, tuple]) -> None: ...


class CellNumberFormat(object):
    def __init__(self):
        self.format: str = 'notSet'

    def set_format(self, string: Optional[str]) -> None: ...


class CellProtection(object):
    def __init__(self):
        self.hidden: bool = ...
        self.locked: bool = ...

    def set_hidden(self, on_off: bool) -> None: ...

    def set_locked(self, on_off: bool) -> None: ...


class CellStyleCopier(object):
    def __init__(self, from_cell: Cell):
        self._style = ...
        self._font: Font = ...
        self._border: Border = ...
        self._fill: Fill = ...
        self._number_format = ...
        self._protection: Protection = ...
        self._alignment: Alignment = ...

    def to_cell(self, cell: Cell) -> None: ...


def get_color_code(color: Union[str, tuple]) -> str: ...


class NoneStyle(object):
    _instance_lock: Lock = ...

    def __init__(self):
        self._font: Font = ...
        self._border: Border = ...
        self._alignment: Alignment = ...
        self._fill: PatternFill = ...
        self._number_format: str = ...
        self._protection: Protection = ...

    def __new__(cls, *args, **kwargs): ...

    def to_cell(self, cell: Cell, replace: bool = True) -> None: ...
