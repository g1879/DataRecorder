# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from typing import Union, Tuple, Any

from openpyxl import load_workbook, Workbook
from openpyxl.cell import Cell
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string


class BaseRecorder(object):
    """记录器的父类"""
    SUPPORTS = ('xlsx', 'csv')

    def __init__(self, path: Union[str, Path], cache_size: int = None) -> None:
        """初始化                                  \n
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件
        """
        self._data = []
        self._before = []
        self._after = []
        self._type = None
        self.path = path
        self.cache_size = cache_size if cache_size is not None else 1000
        self.encoding: str = 'utf-8'
        self.delimiter: str = ','  # csv文件分隔符
        self.quote_char: str = '"'  # csv文件引用符

    def __del__(self) -> None:
        """对象关闭时把剩下的数据写入文件"""
        self.record()

    @property
    def cache_size(self) -> int:
        """返回缓存大小"""
        return self._cache

    @cache_size.setter
    def cache_size(self, cache_size: int) -> None:
        """设置缓存大小                   \n
        :param cache_size: 缓存大小
        :return: None
        """
        if not isinstance(cache_size, int) or cache_size < 0:
            raise TypeError('cache_size值只能是int，且必须>=0')
        self._cache = cache_size

    @property
    def path(self) -> str:
        """返回文件路径"""
        return self._path

    @path.setter
    def path(self, path: Union[str, Path]) -> None:
        """设置文件路径                \n
        :param path: 文件路径
        :return: None
        """
        if isinstance(path, str):
            self._type = path.split('.')[-1]
        elif isinstance(path, Path):
            self._type = path.suffix[1:]
        else:
            raise TypeError(f'参数file_path只能是str或Path，非{type(path)}。')

        if self._type not in self.SUPPORTS:
            raise TypeError(f'只支持{"、".join(self.SUPPORTS)}格式文件。')

        self.record()  # 更换文件前自动记录剩余数据
        self._path = str(path) if isinstance(path, Path) else path

    @property
    def type(self) -> str:
        """返回文件类型"""
        return self._type

    @property
    def data(self) -> list:
        """返回当前保存在缓存的数据"""
        return self._data

    @property
    def before(self) -> Union[list, tuple, str, dict]:
        """返回当前before内容"""
        return self._before

    @property
    def after(self) -> Union[list, tuple, str, dict]:
        """返回当前after内容"""
        return self._after

    def set_before(self, before: Union[list, tuple, str, dict]) -> None:
        """设置在数据前面补充的列                              \n
        :param before: 列表、元组或字符串，为字符串时则补充一列
        :return: None
        """
        self.record()

        if before is None:
            self._before = []
        elif isinstance(before, (list, dict)):
            self._before = before
        elif isinstance(before, tuple):
            self._before = list(before)
        else:
            self._before = [before]

    def set_after(self, after: Union[list, tuple, str, dict]) -> None:
        """设置在数据后面补充的列                                \n
        :param after: 列表、元组或字符串，为字符串时则补充一列
        :return: None
        """
        self.record()

        if after is None:
            self._after = []
        elif isinstance(after, (list, dict)):
            self._after = after
        elif isinstance(after, tuple):
            self._after = list(after)
        else:
            self._after = [after]

    def clear(self) -> None:
        """清空缓存中的数据"""
        self._data = []

    def record(self) -> None:
        """记录数据"""
        # 具体功能由_record()实现，本方法实现自动重试功能
        if not self._data:
            return

        while True:
            try:
                self._record()
                break

            except PermissionError:
                print('\r文件被打开，保存失败，请关闭，程序会自动重试...', end='')

            except Exception as e:
                if self._data:
                    print(f'\n{self._data}\n\n注意！！以上数据未保存')
                if 'Python is likely shutting down' not in str(e):
                    raise
                break

        self._data = []

    def set_head(self, head: Union[list, tuple]) -> None:
        """设置表头。只有 csv 和 xlsx 格式支持设置表头       \n
        :param head: 表头，列表或元组
        :return: None
        """
        if self.type == 'xlsx':
            _set_xlsx_head(self.path, head)

        elif self.type == 'csv':
            _set_csv_head(self.path, head, self.encoding, self.delimiter, self.quote_char)

        else:
            raise TypeError('只能对xlsx和csv文件设置表头。')

    @abstractmethod
    def add_data(self, data):
        pass

    @abstractmethod
    def _record(self):
        pass


def _data_to_list(data: Union[list, tuple, dict],
                  before: Union[list, dict, None] = None,
                  after: Union[list, dict, None] = None,
                  process_content: bool = False) -> list:
    """将传入的数据转换为列表形式                                  \n
    :param data: 要处理的数据
    :param before: 数据前的列
    :param after: 数据后的列
    :param process_content: 是否执行process_content处理数据
    :return: 转变成列表方式的数据
    """
    return_list = []
    if data is not None and not isinstance(data, (list, tuple, dict)):
        data = [data]
    if before is not None and not isinstance(before, (list, tuple, dict)):
        before = [before]
    if after is not None and not isinstance(after, (list, tuple, dict)):
        after = [after]

    for i in (before, data, after):
        if isinstance(i, dict):
            return_list.extend(list(i.values()))
        elif i is None:
            pass
        elif isinstance(i, list):
            return_list.extend(i)
        elif isinstance(i, tuple):
            return_list.extend(list(i))
        else:
            return_list.extend([str(i)])

    return [_process_content(i) for i in return_list] if process_content else return_list


def _data_to_list_or_dict(data: Union[list, tuple, dict],
                          before: Union[list, tuple, dict, None] = None,
                          after: Union[list, tuple, dict, None] = None,
                          process_content: bool = False) -> Union[list, dict]:
    """将传入的数据转换为列表或字典形式，用于记录到txt或json          \n
    :param data: 要处理的数据
    :param before: 数据前的列
    :param after: 数据后的列
    :param process_content: 是否执行process_content处理数据
    :return: 转变成列表或字典形式的数据
    """
    if isinstance(data, (list, tuple)):
        return _data_to_list(data, before, after, process_content)

    elif isinstance(data, dict):
        if isinstance(before, dict):
            data = {**before, **data}

        if isinstance(after, dict):
            data = {**data, **after}

        return data


def _set_csv_head(file_path: str,
                  head: Union[list, tuple],
                  encoding: str = 'utf-8',
                  delimiter: str = ',',
                  quotechar: str = '"') -> None:
    """设置csv文件的表头              \n
    :param file_path: 文件路径
    :param head: 表头列表或元组
    :param encoding: 编码
    :param delimiter: 分隔符
    :param quotechar: 引用符
    :return: None
    """
    from csv import writer
    if Path(file_path).exists():
        with open(file_path, 'r', newline='', encoding=encoding) as f:
            content = "".join(f.readlines()[1:])

        with open(file_path, 'w', newline='', encoding=encoding) as f:
            csv_write = writer(f, delimiter=delimiter, quotechar=quotechar)
            csv_write.writerow(head)

        with open(file_path, 'a+', newline='', encoding=encoding) as f:
            f.write(f'{content}')

    else:
        with open(file_path, 'w', newline='', encoding=encoding) as f:
            csv_write = writer(f, delimiter=delimiter, quotechar=quotechar)
            csv_write.writerow(head)


def _set_xlsx_head(file_path: str, head: Union[list, tuple]) -> None:
    """设置xlsx文件的表头            \n
    :param file_path: 文件路径
    :param head: 表头列表或元组
    :return: None
    """
    wb = load_workbook(file_path) if Path(file_path).exists() else Workbook()
    ws = wb.active

    for key, i in enumerate(head, 1):
        ws.cell(1, key).value = i

    wb.save(file_path)
    wb.close()


def _parse_coord(coord: Union[int, str, list, tuple],
                 col: int = None,
                 disable_type: Union[type, list, tuple] = None) -> Tuple[int, int]:
    """解析坐标，返回坐标tuple                                              \n
    :param coord: 坐标
    :param col: 列号，用于只传入行号的情况
    :param disable_type: 禁用的格式，可以传入单个格式或格式组成的list或tuple
    :return: 坐标tuple（行, 列）
    """
    if disable_type and isinstance(coord, disable_type):
        raise TypeError(f'当前坐标类型不允许为{disable_type}')

    if isinstance(coord, (int, float)):
        if col:
            return int(coord), col
        else:
            raise ValueError('只输入行号时必须同时输入列号')

    if isinstance(coord, str):
        if ',' in coord:  # '3,1'形式
            coord = coord.replace(' ', '').split(',')
        else:  # 'A3'形式
            xy = coordinate_from_string(coord)
            return xy[1], column_index_from_string(xy[0])

    if isinstance(coord, (tuple, list)) and len(coord) == 2:
        return int(coord[0]), int(coord[1])
    else:
        raise ValueError('list或tuple时长度必须为2')


def _process_content(content: Any) -> Union[int, str, float, None]:
    """处理要写入的数据                  \n
    :param content: 未处理的数据内容
    :return: 处理后的数据
    """
    if isinstance(content, (int, str, float, type(None))):
        return content
    elif isinstance(content, Cell):
        return content.value
    else:
        return str(content)
