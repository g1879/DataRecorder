# -*- coding:utf-8 -*-
from csv import reader as csv_reader, writer as csv_writer

from pathlib import Path
from typing import Union


def align_csv(path: Union[str, Path], encoding='utf-8') -> None:
    """补全csv文件，使其每行列数一样多，用于pandas读取时避免出错
    :param path: 要处理的文件路径
    :param encoding: 文件编码
    :return: None
    """
    with open(path, 'r', encoding=encoding) as f:
        reader = csv_reader(f)
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

        writer = csv_writer(open(path, 'w', encoding=encoding, newline=''))
        writer.writerows(lines)
