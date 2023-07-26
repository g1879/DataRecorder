# -*- coding:utf-8 -*-
from copy import copy
from csv import reader as csv_reader, writer as csv_writer

from pathlib import Path
from re import search, sub


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


def align_csv(path, encoding='utf-8', delimiter=',', quotechar='"'):
    """补全csv文件，使其每行列数一样多，用于pandas读取时避免出错
    :param path: 要处理的文件路径
    :param encoding: 文件编码
    :param delimiter: 分隔符
    :param quotechar: 引用符
    :return: None
    """
    with open(path, 'r', encoding=encoding) as f:
        reader = csv_reader(f, delimiter=delimiter, quotechar=quotechar)
        lines = list(reader)
        lines_data = {}
        max_len = 0

        # 把每行列数用字典记录，并找到最长的一行
        for k, i in enumerate(lines):
            line_len = len(i)
            if line_len > max_len:
                max_len = line_len
            lines_data[k] = line_len

        # 把所有行用空值补全到和最长一行一样
        for i in lines_data:
            lines[i].extend([None] * (max_len - lines_data[i]))

        writer = csv_writer(open(path, 'w', encoding=encoding, newline=''), delimiter=delimiter, quotechar=quotechar)
        writer.writerows(lines)


def get_usable_path(path):
    """检查文件或文件夹是否有重名，并返回可以使用的路径
    :param path: 文件或文件夹路径
    :return: 可用的路径，Path对象
    """
    path = Path(path)
    parent = path.parent
    path = parent / make_valid_file_name(path.name)
    name = path.stem if path.is_file() else path.name
    ext = path.suffix if path.is_file() else ''

    first_time = True

    while path.exists():
        r = search(r'(.*)_(\d+)$', name)

        if not r or (r and first_time):
            src_name, num = name, '1'
        else:
            src_name, num = r.group(1), int(r.group(2)) + 1

        name = f'{src_name}_{num}'
        path = parent / f'{name}{ext}'
        first_time = None

    return path


def make_valid_file_name(full_name):
    """获取有效的文件名
    :param full_name: 文件名
    :return: 可用的文件名
    """
    # ----------------去除前后空格----------------
    full_name = full_name.strip()

    # ----------------使总长度不大于255个字符（一个汉字是2个字符）----------------
    r = search(r'(.*)(\.[^.]+$)', full_name)  # 拆分文件名和后缀名
    if r:
        name, ext = r.group(1), r.group(2)
        ext_long = len(ext)
    else:
        name, ext = full_name, ''
        ext_long = 0

    while _get_long(name) > 255 - ext_long:
        name = name[:-1]

    full_name = f'{name}{ext}'

    # ----------------去除不允许存在的字符----------------
    return sub(r'[<>/\\|:*?\n]', ' ', full_name)


def _get_long(txt) -> int:
    """返回字符串中字符个数（一个汉字是2个字符）
    :param txt: 字符串
    :return: 字符个数
    """
    txt_len = len(txt)
    return int((len(txt.encode('utf-8')) - txt_len) / 2 + txt_len)
