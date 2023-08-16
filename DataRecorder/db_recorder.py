# -*- coding:utf-8 -*-
from sqlite3 import connect
from time import sleep

from .base import BaseRecorder
from .setter import DBSetter
from .tools import data_to_list_or_dict


class DBRecorder(BaseRecorder):
    SUPPORTS = ('db',)

    def __init__(self, path=None, cache_size=None, table=None):
        """用于存储数据到sqlite的工具
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        :param table: 默认表名
        """
        self._conn = None
        self._cur = None
        self._type = 'db'
        super().__init__(None, cache_size)
        if path:
            self.set.path(path, table)

    @property
    def set(self):
        """返回用于设置属性的对象"""
        if self._setter is None:
            self._setter = DBSetter(self)
        return self._setter

    def __del__(self):
        """重写父类方法"""
        super().__del__()
        self._close_connection()

    def add_data(self, data, table=None):
        """添加数据
        :param data: 可以是一维或二维数据，dict格式可向对应列填写数据，其余格式按顺序从左到右填入各列
        :param table: 数据要插入的表名称
        :return: None
        """
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.1)

        table = table or self.table
        if not isinstance(table, str):
            raise RuntimeError('未指定数据库表名。')

        if table not in self._data:
            self._data[table] = []

        if not isinstance(data, (list, tuple, dict)):
            data = (data,)

        # 一维数组
        if (isinstance(data, (list, tuple)) and not isinstance(data[0], (list, tuple, dict))) or isinstance(data, dict):
            self._data_count += 1
            self._data[table].append(data)

        else:  # 二维数组
            self._data[table].extend(data)
            self._data_count += len(data)

        if 0 < self.cache_size <= self._data_count:
            self.record()

    def run_sql(self, sql, single=True, commit=False):
        """执行sql语句并返回结果
        :param sql: sql语句
        :param single: 是否只获取一个结果
        :param commit: 是否提交到数据库
        :return: 查找到的结果，没有结果时返回None
        """
        self._cur.execute(sql)
        r = self._cur.fetchone() if single else self._cur.fetchall()
        if commit:
            self._conn.commit()
        return r

    def _connect(self):
        """链接数据库"""
        self._conn = connect(self.path)
        self._cur = self._conn.cursor()

    def _close_connection(self):
        """关闭数据库 """
        if self._conn is not None:
            self._cur.close()
            self._conn.close()

    def _to_database(self, data_list, table, tables):
        """把数据批量写入指定数据表
        :param data_list: 要写入的数据组成的列表
        :param table: 要写入数据的数据表名称
        :param tables: 数据库中数据表和列信息
        :return: None
        """
        if isinstance(data_list[0], dict):  # 检查是否要新增列
            keys = data_list[0].keys()
            for key in keys:
                if key not in tables[table]:
                    sql = f'ALTER TABLE {table} ADD COLUMN {key}'
                    self._cur.execute(sql)
                    tables[table].append(key)

            question_masks = ','.join('?' * len(keys))
            keys_txt = ','.join(keys)
            values = [list(i.values()) for i in data_list]
            sql = f'INSERT INTO {table} ({keys_txt}) values ({question_masks})'

        else:
            question_masks = ','.join('?' * len(tables[table]))
            values = data_list
            sql = f'INSERT INTO {table} values ({question_masks})'

        self._cur.executemany(sql, values)

    def _record(self):
        """保存数据到sqlite"""
        # 获取所有表名和列名
        self._cur.execute("select name from sqlite_master where type='table'")
        tables = {}
        for table in self._cur.fetchall():
            self._cur.execute(f"PRAGMA table_info({table[0]})")
            tables[table[0]] = [i[1] for i in self._cur.fetchall()]

        for table, data in self._data.items():
            data_list = []
            if isinstance(data[0], dict):
                curr_keys = data[0].keys()
            else:
                curr_keys = len(data[0])

            for d in data:
                if isinstance(d, dict):
                    tmp_keys = d.keys()
                    d = data_to_list_or_dict(self, d)
                    if table not in tables:
                        keys = d.keys()
                        self._cur.execute(f"CREATE TABLE {table} ({','.join(keys)})")
                        tables[table] = tuple(keys)

                else:
                    if table not in tables:
                        raise TypeError('新建表格首次须接收数据需为dict格式。')
                    tmp_keys = len(d)
                    d = self._data_to_list(d, len(tables[table]))

                if tmp_keys != curr_keys:
                    self._to_database(data_list, table, tables)
                    curr_keys = tmp_keys
                    data_list = []

                data_list.append(d)

            if data_list:
                self._to_database(data_list, table, tables)

        self._conn.commit()
