# -*- coding:utf-8 -*-
from copy import copy

from openpyxl.styles import Font, Side


class CellStyle(object):
    def __init__(self):
        self._style = None
        self._font = None
        self._border = None
        self._fill = None
        self._number_format = None
        self._protection = None
        self._alignment = None

    @property
    def font(self):
        if self._font is None:
            self._font = CellFont()
        return self._font

    def to_cell(self, cell):
        """把当前样式复制到目标单元格"""
        if self._font:
            args = ('name', 'size', 'charset', 'underline', 'color', 'scheme', 'vertAlign',
                    'bold', 'italic', 'strike', 'outline', 'shadow', 'condense', 'extend')
            d = {}
            for arg in args:
                tmp = getattr(self._font, arg)
                d[arg] = getattr(cell.font, arg) if tmp == 'notSet' else tmp

            cell.font = Font(name=d['name'], sz=d['size'], b=d['bold'], i=d['italic'], charset=d['charset'],
                             u=d['underline'], strike=d['strike'], color=d['color'], scheme=d['scheme'],
                             family=cell.font.family, vertAlign=d['vertAlign'], outline=d['outline'],
                             shadow=d['shadow'], condense=d['condense'], extend=d['extend'])

        if self._alignment:
            cell.alignment = self._alignment
        if self._border:
            cell.border = self._border
        if self._fill:
            cell.fill = self._fill
        if self._number_format:
            cell.number_format = self._number_format
        if self._protection:
            cell.protection = self._protection


class CellFont(object):
    def __init__(self):
        self.name = 'notSet'
        self.charset = 'notSet'
        self.size = 'notSet'
        self.bold = 'notSet'
        self.italic = 'notSet'
        self.strike = 'notSet'
        self.outline = 'notSet'
        self.shadow = 'notSet'
        self.condense = 'notSet'
        self.extend = 'notSet'
        self.underline = 'notSet'
        self.vertAlign = 'notSet'
        self.color = 'notSet'
        self.scheme = 'notSet'

    def set_name(self, name):
        """设置字体
        :param name: 字体名称
        :return: None
        """
        self.name = name

    def set_charset(self, charset):
        """设置编码
        :param charset: 字体编码，int格式
        :return: None
        """
        if not isinstance(charset, int):
            raise TypeError('charset参数只能接收int类型。')
        self.charset = charset

    def set_size(self, size):
        """设置字体大小
        :param size: 字体大小
        :return: None
        """
        self.size = size

    def set_bold(self, on_off):
        """设置是否加粗
        :param on_off: bool表示开关
        :return: None
        """
        self.bold = on_off

    def set_italic(self, on_off):
        """设置是否斜体
        :param on_off: bool表示开关
        :return: None
        """
        self.italic = on_off

    def set_strike(self, on_off):
        """设置是否有删除线
        :param on_off: bool表示开关
        :return: None
        """
        self.strike = on_off

    def set_outline(self, on_off):
        """设置outline
        :param on_off: bool表示开关
        :return: None
        """
        self.outline = on_off

    def set_shadow(self, on_off):
        """设置是否有阴影
        :param on_off: bool表示开关
        :return: None
        """
        self.shadow = on_off

    def set_condense(self, on_off):
        """设置condense
        :param on_off: bool表示开关
        :return: None
        """
        self.condense = on_off

    def set_extend(self, on_off):
        """设置extend
        :param on_off: bool表示开关
        :return: None
        """
        self.extend = on_off

    def set_color(self, color):
        """设置字体颜色
        :param color: 字体演示字符串，如`FFF000`
        :return: None
        """
        self.color = color

    def set_underline(self, option):
        """设置下划线
        :param option: 下划线类型，可选 'single', 'double', 'singleAccounting', 'doubleAccounting'，False为取消
        :return: None
        """
        if option not in ('single', 'double', 'singleAccounting', 'doubleAccounting', None):
            raise ValueError(
                "option参数只能是'single', 'double', 'singleAccounting', 'doubleAccounting', None其中之一。")
        self.underline = option

    def set_vertAlign(self, option):
        """设置居中？
        :param option: 可选 'superscript', 'subscript', 'baseline'
        :return: None
        """
        if option not in ('superscript', 'subscript', 'baseline', None):
            raise ValueError(
                "option参数只能是'single', 'double', 'singleAccounting', 'doubleAccounting', None其中之一。")
        self.vertAlign = option

    def set_scheme(self, option):
        """设置scheme
        :param option: 可选 'major', 'minor'
        :return: None
        """
        if option not in ('single', 'double', 'singleAccounting', 'doubleAccounting', False):
            raise ValueError("option参数只能是'major', 'minor', False其中之一。")
        self.scheme = option


class CellBorder(object):
    line_style = ('dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'medium', 'mediumDashDot',
                  'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin')

    def __init__(self):
        self.start = 'notSet'
        self.end = 'notSet'
        self.left = 'notSet'
        self.right = 'notSet'
        self.top = 'notSet'
        self.bottom = 'notSet'
        self.diagonal = 'notSet'
        self.vertical = 'notSet'
        self.horizontal = 'notSet'
        self.horizontal = 'notSet'
        self.outline = 'notSet'
        self.diagonalUp = 'notSet'
        self.diagonalDown = 'notSet'

    def set_start(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.start = Side(style=style, color=color)

    def set_end(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.end = Side(style=style, color=color)

    def set_left(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.left = Side(style=style, color=color)

    def set_right(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.right = Side(style=style, color=color)

    def set_top(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.top = Side(style=style, color=color)

    def set_bottom(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.bottom = Side(style=style, color=color)

    def set_diagonal(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.diagonal = Side(style=style, color=color)

    def set_vertical(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.vertical = Side(style=style, color=color)

    def set_horizontal(self, style, color):
        """设置
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'
        :param color: 颜色
        :return: None
        """
        if style not in self.line_style:
            raise ValueError(f'style参数只能是{self.line_style}之一。')
        self.horizontal = Side(style=style, color=color)

    def set_outline(self, on_off):
        """
        :param on_off: bool表示开关
        :return: None
        """
        self.outline = on_off

    def set_diagonalDown(self, on_off):
        """
        :param on_off: bool表示开关
        :return: None
        """
        self.diagonalDown = on_off

    def set_diagonalUp(self, on_off):
        """
        :param on_off: bool表示开关
        :return: None
        """
        self.diagonalUp = on_off
