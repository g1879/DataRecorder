# -*- coding:utf-8 -*-
from pathlib import Path
from sqlite3 import connect
from time import sleep
from typing import Union, Any

from DataRecorder import Recorder


class DBRecorder(Recorder):
    SUPPORTS = ('db',)

    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = None,
                 table: str = None):
        """
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        :param table: 默认表名
        """
        super().__init__(path, cache_size)
        self.table = table

    @property
    def table(self) -> str:
        """返回默认表名"""
        return self._table

    @table.setter
    def table(self, name: str) -> None:
        """设置默认表名"""
        self._table = name

    def set_table(self, table: str) -> None:
        """设置默认表名
        :param table: 表名
        :return: None
        """
        self._table = table

    def add_data(self, data: Any, table: str = None) -> None:
        """添加数据                                                                            \n
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

        if 0 < self.cache_size <= len(self._data):
            self.record()

    def _record(self) -> None:
        """记录数据"""
        if self.type == 'db':
            self._to_sqlite()

    def _to_sqlite(self) -> None:
        """保存数据到sqlite"""
        # 每个数据格式为(表名，一维或二维数据)
        new = not Path(self.path).exists()

        conn = connect(self.path)
        cur = conn.cursor()
        if new:
            data = self.data[0]
            table = data[0]
            data = data[1]
            _create_table(cur, table, data)

        # 获取所有表名和列名
        cur.execute("select name from sqlite_master where type='table'")
        tables = {}
        for table in cur.fetchall():
            cur.execute(f"PRAGMA table_info({table[0]})")
            tables[table[0]] = [i[1] for i in cur.fetchall()]

        # 添加数据
        for data in self._data:
            table = data[0]
            data = data[1]

            if table not in tables:
                name, cols = _create_table(cur, table, data)
                tables[table] = cols

            now_data = (data,) if not isinstance(data[0], (list, tuple, dict)) else data

            for d in now_data:
                if isinstance(d, dict):
                    d = self._data_to_list_or_dict(d)
                    keys = d.keys()

                    for key in keys:  # 检查是否要新增列
                        if key not in tables[table]:
                            sql = f'ALTER TABLE {table} ADD COLUMN {key}'
                            cur.execute(sql)
                            tables[table].append(key)

                    keys_txt = ','.join(keys)
                    values = ','.join([str(i) if not isinstance(i, str) else f'"{i}"' for i in d.values()])

                else:
                    d = self._data_to_list(d)
                    long = len(d)
                    if long > len(tables[table]):
                        raise RuntimeError('数据个数大于列数。')

                    keys_txt = ','.join(tables[table][:long])
                    values = ','.join([str(i) if not isinstance(i, str) else f'"{i}"' for i in d])

                sql = f'INSERT INTO {table} ({keys_txt}) values ({values})'
                cur.execute(sql)

        conn.commit()
        cur.close()
        conn.close()


def _create_table(cursor, table_name: str, data: dict) -> tuple:
    """创建表格                                 \n
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
