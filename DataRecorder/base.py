# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from threading import Lock
from time import sleep

from .setter import OriginalSetter, BaseSetter
from .tools import get_usable_path


class OriginalRecorder(object):
    """记录器的基类"""
    SUPPORTS = ('any',)

    def __init__(self, path=None, cache_size=None):
        """
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        self._data = []
        self._path = None
        self._type = None
        self._lock = Lock()
        self._pause_add = False  # 文件写入时暂停接收输入
        self._pause_write = False  # 标记文件正在被一个线程写入
        self.show_msg = True
        self._setter = None

        if path:
            self.set.path(path)
        self._cache = cache_size if cache_size is not None else 1000

    def __del__(self):
        """对象关闭时把剩下的数据写入文件"""
        self.record()

    @property
    def set(self):
        """返回用于设置属性的对象"""
        if self._setter is None:
            self._setter = OriginalSetter(self)
        return self._setter

    @property
    def cache_size(self):
        """返回缓存大小"""
        return self._cache

    @property
    def path(self):
        """返回文件路径"""
        return self._path

    @property
    def type(self):
        """返回文件类型"""
        return self._type

    @property
    def data(self):
        """返回当前保存在缓存的数据"""
        return self._data

    def record(self, new_path=None):
        """记录数据，可保存到新文件
        :param new_path: 文件另存为的路径，会保存新文件
        :return: 成功返回文件路径，失败返回未保存的数据
        """
        # 具体功能由_record()实现，本方法实现自动重试及另存文件功能
        original_path = return_path = self._path
        return_data = None
        if new_path:
            new_path = str(get_usable_path(new_path))
            return_path = self._path = new_path

            if Path(original_path).exists():
                from shutil import copy
                copy(original_path, self._path)

        if not self._data:
            return return_path

        if not self._path:
            raise ValueError('保存路径为空。')

        with self._lock:
            self._pause_add = True  # 写入文件前暂缓接收数据
            if self.show_msg:
                print(f'{self.path} 开始写入文件，切勿关闭进程')

            Path(self.path).parent.mkdir(parents=True, exist_ok=True)
            while True:
                try:
                    while self._pause_write:  # 等待其它线程写入结束
                        sleep(.1)

                    self._pause_write = True
                    self._record()
                    break

                except PermissionError:
                    if self.show_msg:
                        print('\r文件被打开，保存失败，请关闭，程序会自动重试...', end='')

                except Exception as e:
                    if self._data:
                        if self.show_msg:
                            print(f'{"=" * 30}\n{self._data}\n\n自动写入失败，以上数据未保存。\n'
                                  f'错误信息：{e}\n'
                                  f'提醒：请显式调用record()保存数据。\n{"=" * 30}')
                            raise
                        return_data = self._data.copy()
                    break

                finally:
                    self._pause_write = False

                sleep(.3)

            if new_path:
                self._path = original_path

            if self.show_msg and not return_data:
                print(f'{self.path} 写入文件结束')
            self._data = []
            self._pause_add = False

        return return_data if return_data else return_path

    def clear(self):
        """清空缓存中的数据"""
        self._data = []

    @abstractmethod
    def add_data(self, data):
        pass

    @abstractmethod
    def _record(self):
        pass


class BaseRecorder(OriginalRecorder):
    """Recorder和Filler的父类"""
    SUPPORTS = ('xlsx', 'csv')

    def __init__(self, path=None, cache_size=None):
        """
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        super().__init__(path, cache_size)
        self._before = []
        self._after = []

        self._encoding = 'utf-8'
        self._table = None

    @property
    def set(self):
        """返回用于设置属性的对象"""
        if self._setter is None:
            self._setter = BaseSetter(self)
        return self._setter

    @property
    def before(self):
        """返回当前before内容"""
        return self._before

    @property
    def after(self):
        """返回当前after内容"""
        return self._after

    @property
    def table(self):
        """返回默认表名"""
        return self._table

    @property
    def encoding(self):
        """返回编码格式"""
        return self._encoding

    @abstractmethod
    def add_data(self, data):
        pass

    @abstractmethod
    def _record(self):
        pass

    def _data_to_list(self, data):
        """将传入的数据转换为列表形式，添加前后列数据
        :param data: 要处理的数据
        :return: 转变成列表方式的数据
        """
        return_list = []
        if data is not None and not isinstance(data, (list, tuple, dict)):
            data = [data]

        for i in (self.before, data, self.after):
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

        return return_list
