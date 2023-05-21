# -*- coding:utf-8 -*-
from copy import copy


class CellStyle(object):
    def __init__(self, from_cell=None):
        if from_cell:
            self._style = copy(from_cell._style)
            self._font = copy(from_cell.font)
            self._border = copy(from_cell.border)
            self._fill = copy(from_cell.fill)
            self._number_format = copy(from_cell.number_format)
            self._protection = copy(from_cell.protection)
            self._alignment = copy(from_cell.alignment)
        else:
            self._style = None
            self._font = None
            self._border = None
            self._fill = None
            self._number_format = None
            self._protection = None
            self._alignment = None

    def set_alignment(self, alignment):
        self._alignment = alignment

    def set_font(self, font):
        self._font = font

    def set_border(self, border):
        self._border = border

    def set_fill(self, fill):
        self._fill = fill

    def set_number_format(self, number_format):
        self._number_format = number_format

    def set_protection(self, protection):
        self._protection = protection

    def to_cell(self, cell):
        """把当前样式复制到目标单元格"""
        if self._style:
            cell._style = self._style
        if self._alignment:
            cell.alignment = self._alignment
        if self._font:
            cell.font = self._font
        if self._border:
            cell.border = self._border
        if self._fill:
            cell.fill = self._fill
        if self._number_format:
            cell.number_format = self._number_format
        if self._protection:
            cell.protection = self._protection
