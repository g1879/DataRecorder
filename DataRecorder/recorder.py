# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union

from openpyxl import Workbook, load_workbook

from .base import BaseRecorder, _data_to_list, _data_to_list_or_dict


class Recorder(BaseRecorder):
    """用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销。
    退出时能自动记录数据（xlsx格式除外），避免因异常丢失。
    """
    SUPPORTS = ('xlsx', 'csv', 'json', 'txt')

    def add_data(self, data: Union[list, tuple, dict, int, float, str]) -> None:
        """添加数据，可一次添加多条数据                  \n
        :param data: 插入的数据，元组或列表
        :return: None
        """
        if not isinstance(data, (list, tuple, dict)):
            data = (data,)
        if data and (isinstance(data, (list, tuple)) and not isinstance(data[0], (list, tuple, dict))) \
                or isinstance(data, dict):
            self._data.append(data)
        else:
            self._data.extend(data)

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def _record(self) -> None:
        """记录数据"""
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)

        if self.type == 'xlsx':
            _record_to_xlsx(self.path, self._data, self._before, self._after)
        elif self.type == 'csv':
            _record_to_csv(self.path, self._data, self._before, self._after, self.encoding, self.delimiter,
                           self.quote_char)
        elif self.type == 'txt':
            _record_to_txt(self.path, self._data, self._before, self._after, self.encoding)
        elif self.type == 'json':
            _record_to_json(self.path, self._data, self._before, self._after, self.encoding)


def _get_title(data: Union[list, dict],
               before: Union[list, dict, None] = None,
               after: Union[list, dict, None] = None) -> Union[list, None]:
    """获取表头列表                  \n
    :param data: 数据列表或字典
    :param before: 数据前的列
    :param after: 数据后的列
    :return: 表头列表
    """
    if isinstance(data, (tuple, list)):
        return None

    return_list = []
    for i in (before, data, after):
        if isinstance(i, dict):
            return_list.extend(list(i))
        elif i is None:
            pass
        elif isinstance(i, list):
            return_list.extend(['' for _ in range(len(i))])
        else:
            return_list.extend([''])

    return return_list


def _record_to_xlsx(file_path: str,
                    data: list,
                    before: Union[list, tuple, dict] = None,
                    after: Union[list, tuple, dict] = None) -> None:
    """记录数据到xlsx文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :return: None
    """
    if Path(file_path).exists():
        wb = load_workbook(file_path)
        ws = wb.active

    else:
        wb = Workbook(write_only=True)
        ws = wb.create_sheet()

        title = _get_title(data[0], before, after)
        if title is not None:
            ws.append(title)

    for i in data:
        ws.append(_data_to_list(i, before, after, True))

    wb.save(file_path)
    wb.close()


def _record_to_csv(file_path: str,
                   data: list,
                   before: Union[list, dict] = None,
                   after: Union[list, dict] = None,
                   encoding: str = 'utf-8',
                   delimiter: str = ',',
                   quotechar: str = '"') -> None:
    """记录数据到csv文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :param encoding: 编码
    :param delimiter: 分隔符
    :param quotechar: 引用符
    :return: None
    """
    from csv import writer
    title = _get_title(data[0], before, after) if not Path(file_path).exists() else None

    with open(file_path, 'a+', newline='', encoding=encoding) as f:
        csv_write = writer(f, delimiter=delimiter, quotechar=quotechar)
        if title:
            csv_write.writerow(title)
        for i in data:
            csv_write.writerow(_data_to_list(i, before, after, True))


def _record_to_txt(file_path: str,
                   data: list,
                   before: Union[list, dict] = None,
                   after: Union[list, dict] = None,
                   encoding: str = 'utf-8') -> None:
    """记录数据到txt文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :param encoding: 编码
    :return: None
    """
    with open(file_path, 'a+', encoding=encoding) as f:
        all_data = [f'{_data_to_list_or_dict(i, before, after, True)}\n' for i in data]
        f.write(''.join(all_data))


def _record_to_json(file_path: str,
                    data: list,
                    before: Union[list, dict] = None,
                    after: Union[list, dict] = None,
                    encoding: str = 'utf-8') -> None:
    """记录数据到json文件            \n
    :param file_path: 文件路径
    :param data: 要记录的数据
    :param before: 数据前面的列
    :param after: 数据后面的列
    :param encoding: 编码
    :return: None
    """
    from json import load, dump

    if Path(file_path).exists():
        with open(file_path, 'r', encoding=encoding) as f:
            json_data = load(f)

        for i in data:
            json_data.append(_data_to_list_or_dict(i, before, after, True))

    else:
        json_data = [_data_to_list_or_dict(i, before, after, True) for i in data]

    with open(file_path, 'w', encoding=encoding) as f:
        dump(json_data, f)
