# -*- coding:utf-8 -*-
from pathlib import Path
from time import sleep
from typing import Union, Any

from openpyxl import Workbook, load_workbook

from .base import BaseRecorder, _ok_list


class Recorder(BaseRecorder):
    """用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销。
    退出时能自动记录数据（xlsx格式除外），避免因异常丢失。
    """
    SUPPORTS = ('any',)

    def add_data(self, data: Any) -> None:
        """添加数据，可一次添加多条数据                  \n
        :param data: 插入的数据，元组或列表
        :return: None
        """
        while self._pause_add:
            sleep(.1)

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
            self._to_xlsx()
        elif self.type == 'csv':
            self._to_csv()
        elif self.type == 'json':
            self._to_json()
        else:
            self._to_txt()

    def _to_xlsx(self) -> None:
        """记录数据到xlsx文件"""
        if Path(self.path).exists():
            wb = load_workbook(self.path)
            ws = wb.active

        else:
            wb = Workbook(write_only=True)
            ws = wb.create_sheet()
            title = _get_title(self._data[0], self._before, self._after)
            if title is not None:
                ws.append(_ok_list(title, True))

        for i in self._data:
            ws.append(_ok_list(self._data_to_list(i), True))

        wb.save(self.path)
        wb.close()

    def _to_csv(self) -> None:
        """记录数据到csv文件"""
        from csv import writer
        title = _get_title(self._data[0], self._before, self._after) if not Path(self.path).exists() else None

        with open(self.path, 'a+', newline='', encoding=self.encoding) as f:
            csv_write = writer(f, delimiter=self.delimiter, quotechar=self.quote_char)
            if title:
                csv_write.writerow(_ok_list(title))
            for i in self._data:
                csv_write.writerow(_ok_list(self._data_to_list(i)))

    def _to_txt(self) -> None:
        """记录数据到txt文件"""
        with open(self.path, 'a+', encoding=self.encoding) as f:
            all_data = [' '.join(_ok_list(self._data_to_list_or_dict(i), as_str=True)) for i in self._data]
            f.write('\n'.join(all_data) + '\n')

    def _to_json(self) -> None:
        """记录数据到json文件"""
        from json import load, dump
        if Path(self.path).exists():
            with open(self.path, 'r', encoding=self.encoding) as f:
                json_data = load(f)

            for i in self._data:
                json_data.append(_ok_list(self._data_to_list_or_dict(i)))

        else:
            json_data = [_ok_list(self._data_to_list_or_dict(i)) for i in self._data]

        with open(self.path, 'w', encoding=self.encoding) as f:
            dump(json_data, f)

    def _data_to_list_or_dict(self, data: Union[list, tuple, dict]) -> Union[list, dict]:
        """将传入的数据转换为列表或字典形式，用于记录到txt或json          \n
        :param data: 要处理的数据
        :return: 转变成列表或字典形式的数据
        """
        if isinstance(data, (list, tuple)):
            return self._data_to_list(data)

        elif isinstance(data, dict):
            if isinstance(self.before, dict):
                data = {**self.before, **data}

            if isinstance(self.after, dict):
                data = {**data, **self.after}

            return data


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
