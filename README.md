# 简介

本库用于记录数据表格数据到文件。

使用方便，代码简洁。

可以缓存数据到一定数量再一次写入，减少文件读写次数，降低开销。 可以在程序崩溃时自动保存数据（除 xlsx 格式外），避免数据丢失。 可以以表格某些列作为关键字，获取或处理数据后填回表格。

支持 xlsx、csv、json、txt 格式。

# 简单演示

Recorder 演示

```python
file = r'results.csv'  # 用于记录数据的文件
r = Recorder(file, 50)  # 50表示每50条记录写入一次文件
for _ in range(100):  # 产生100条数据
    data = (1, 2, 3, 4)
    r.add_data(data)  # 插入一条数据（也可一次插入多条）
# 程序结束时自动保存文件
```

Filler 演示

```python
file = r'results.csv'
f = Filler(file, key_cols='A', sign_col='B')
# =============方法一=============
for key in f.keys:  # 所有未填充行的key列
    data = do_sth(key, *args)  # 处理数据的方法，第一个参数必须是接收key值
    f.add_data(data)

# =============方法二=============
f.fill(do_sth, *args)  # 调用处理数据方法，自动填充数据
```

# 使用方法

## 安装

```
pip install DataRecorder
```

## 导入

```python
from DataRecorder import Recorder  # 记录器
from DataRecorder import Filler  # 填充器
```

## Recorder 类

Recorder 用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销。退出时能自动记录数据（除 xlsx 格式外），避免因异常丢失。支持 xlsx、csv、json、txt 格式。

### 创建 Recorder 对象

```
r = Recorder(path, 50)  # 传入文件路径，及缓存条数
```

### Recorder 类属性

```python
r.path  # 文件路径
r.cache_size  # 缓存的数据条数
r.type  # 文件类型
```

### Recorder 类方法

```python
r.add_data(data)  # 插入一条或多条数据
r.record()  # 主动保存数据
r.clear()  # 清空缓存中的数据
r.set_before(before)  # 设置在数据前面补充的列
r.set_after(after)  # 设置在数据后面补充的列
r.set_head(head)  # 设置表头。只有 csv 和 xlsx 格式支持设置表头
```

### Tips

- add_data() 可以接收 str、int、float、list、tuple、dict 等类型数据
- add_data() 也可以接收这些类型组成的列表，一次插入多条数据
- 除 xlsx 格式外，其它格式在程序退出或崩溃时会自动记录缓存中的数据
- 进行采集时，经常除了插入当前采集的数据，还要在这些数据前面或后面插入固定的数据列，可以用set_before() 和 set_after() 指定这些列，可以接收 str、int、float、list、tuple、dict 等类型数据
- 如果是新文件且传入的数据是 dict 格式，会自动生成表头
- 指定保存文件的路径不必已经存在，会自动创建
