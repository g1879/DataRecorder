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
        self._data = None
        self._path = None
        self._type = None
        self._lock = Lock()
        self._pause_add = False  # 文件写入时暂停接收输入
        self._pause_write = False  # 标记文件正在被一个线程写入
        self.show_msg = True
        self._setter = None
        self._data_count = 0  # 已缓存数据的条数

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
        :return: 文件路径
        """
        # 具体功能由_record()实现，本方法实现自动重试及另存文件功能
        original_path = return_path = self._path
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
                print(f'{self.path} 开始写入文件，切勿关闭进程。')

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
                        print('\r文件被打开，保存失败，请关闭，程序会自动重试。', end='')

                except Exception as e:
                    try:
                        with open('failed_data.txt', 'a+', encoding='utf-8') as f:
                            f.write(str(self.data) + '\n')
                        print('保存失败的数据已保存到failed_data.txt。')
                    except:
                        raise e
                    raise

                finally:
                    self._pause_write = False

                sleep(.3)

            if new_path:
                self._path = original_path

            if self.show_msg:
                print(f'{self.path} 写入文件结束。')
            self.clear()
            self._data_count = 0
            self._pause_add = False

        return return_path

    def clear(self):
        """清空缓存中的数据"""
        if self._data:
            self._data.clear()

    @abstractmethod
    def add_data(self, data):
        pass

    @abstractmethod
    def _record(self):
        pass

    # ---------------即将废弃--------------------
    def set_path(self, path):
        """设置文件路径
        :param path: 文件路径
        :return: None
        """
        self.set.path(path)


class BaseRecorder(OriginalRecorder):
    """Recorder、Filler和DBRecorder的父类"""
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
    def add_data(self, data, table=None):
        pass

    @abstractmethod
    def _record(self):
        pass

    # ---------------即将废弃--------------------
    def set_before(self, before):
        """设置before
        :param before: before
        :return: None
        """
        self.set.before(before)

    def set_after(self, after):
        """设置after
        :param after: after
        :return: None
        """
        self.set.after(after)
