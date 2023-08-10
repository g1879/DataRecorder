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
        super().__init__(None, cache_size)
        if path:
            self.set.path(path, table)
        self._type = 'db'

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

        if not isinstance(data, (list, tuple)):
            data = (data,)

        table = table or self.table
        if table is None:
            raise RuntimeError('未指定数据库表名。')

        self._data.append((table, data))
        self._data_count += len(data[0]) if isinstance(data[0], (list, tuple, dict)) else 1

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

    def _record(self):
        """记录数据"""
        if self.type == 'db':
            self._to_sqlite()

    def _to_sqlite(self):
        """保存数据到sqlite"""
        # 获取所有表名和列名
        self._cur.execute("select name from sqlite_master where type='table'")
        tables = {}
        for table in self._cur.fetchall():
            self._cur.execute(f"PRAGMA table_info({table[0]})")
            tables[table[0]] = [i[1] for i in self._cur.fetchall()]

        # 添加数据，每个数据格式为(表名，一维或二维数据)
        for data in self._data:
            table = data[0]
            data = data[1]  # 一维或二维数组

            if table not in tables:
                d = data[0][0] if isinstance(data[0], (list, tuple)) else data[0]  # 获取第一个数据
                name, cols = _create_table(self._cur, table, d)
                tables[table] = cols

            now_data = (data,) if not isinstance(data[0], (list, tuple, dict)) else data

            for d in now_data:
                if isinstance(d, dict):
                    d = data_to_list_or_dict(self, d)
                    question_masks = ','.join('?' * len(d))
                    keys = d.keys()

                    for key in keys:  # 检查是否要新增列
                        if key not in tables[table]:
                            sql = f'ALTER TABLE {table} ADD COLUMN {key}'
                            self._cur.execute(sql)
                            tables[table].append(key)

                    keys_txt = ','.join(keys)
                    values = list(d.values())
                    sql = f'INSERT INTO {table} ({keys_txt}) values ({question_masks})'

                else:
                    d = self._data_to_list(d)
                    long = len(d)
                    cols_num = len(tables[table])
                    if long > cols_num:
                        raise RuntimeError('数据个数大于列数。')
                    d.extend([None] * (cols_num - long))
                    question_masks = ','.join('?' * cols_num)

                    values = d
                    sql = f'INSERT INTO {table} values ({question_masks})'

                values = [str(i) if i is not None and not isinstance(i, (str, float, int, bool)) else i for i in values]
                self._cur.execute(sql, values)

        self._conn.commit()


def _create_table(cursor, table_name: str, data: dict) -> tuple:
    """创建表格
    :param cursor: 数据库游标对象
    :param table_name: 表名称
    :param data: 要添加的数据
    :return: 表名和各列组成的元组
    """
    if not isinstance(data, dict):
        raise TypeError('新建表格须接收dict格式数据。')

    titles_txt = ','.join(data.keys())
    cursor.execute(f'CREATE TABLE {table_name} ({titles_txt})')

    return table_name, data.keys()
