# 简介

本库是一个使用方便，代码简洁的小工具，用于记录数据到文件。

## 背景

在网上爬取数据的时候，常常需要把数据保存到文件，频繁的开关文件会影响效率，而如果等等爬取结束再写入，会有因异常而导致数据丢失的风险。因此写了一个小工具，须要保存数据时只要把数据扔进去，它会把数据缓存到一定数量再一次写入，而且在程序崩溃时能自动保存数据，保证数据的可靠性。

后来增加了一个功能，从已有文件中读取关键字，爬取数据后自动回填。这样对断点续爬提供了良好的支持，封装了常用的功能，减轻了编码量，让程序员可把更多精力放在业务逻辑。

后来又增加了一个功能，可以把二维数据一次填入 csv 或 xlsx 文件指定左上角坐标的区域中。

## 特性

- 可以缓存数据到一定数量再一次写入，减少文件读写次数，降低开销。
- 可以在程序崩溃时自动保存或显示剩余数据，避免数据丢失。
- 可以以表格某些列作为关键字，获取或处理数据后填回表格。
- 可直接接收 openpyxl 的 Cell 对象记录数据。
- 支持 xlsx、csv、json、txt 格式。

# 简单演示

Recorder 演示

```python
file = 'results.csv'  # 用于记录数据的文件
r = Recorder(file, 50)  # 50表示每50条记录写入一次文件
for _ in range(100):  # 产生100条数据
    data = (1, 2, 3, 4)
    r.add_data(data)  # 插入一条数据（也可一次插入多条）
# 程序结束时自动保存文件
```

Filler 演示

```python
file = 'results.csv'
f = Filler(file, key_cols='A', sign_col='B')
# =============方法一=============
for key in f.keys:  # 所有未填充行的key列
    data = do_sth(key, *args)  # 处理数据的方法，第一个参数必须是接收key值
    f.add_data(data)

# =============方法二=============
f.fill(do_sth, *args)  # 调用处理数据方法，自动填充数据
```

MapGun 演示

```python
file = 'results.csv'
m = MapGun(file)
data = ((1, 2),
        (3, 4))
m.add_data(data, 'c4')  # 把二维数据填入以 c4 为左上角的区域中
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

Recorder 用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销。退出时能自动记录数据，避免因异常丢失。支持 xlsx、csv、json、txt 格式。

### 创建 Recorder 对象

```python
r = Recorder(path, cache_size)  # 传入文件路径，缓存条数
```

### Recorder 类属性

```python
r.path  # 文件路径
r.cache_size  # 缓存的数据条数
r.data  # 返回当前保存的数据
r.type  # 文件类型
r.delimiter  # csv文件分隔符
r.quote_char  # csv文件引用符
r.before  #  补充在前面的列数据
r.after  #  补充在后面的列数据
```

### Recorder 类方法

```python
r.add_data(data)  # 插入一条或多条数据
r.record(new_path)  # 主动保存数据，可指定另存为的路径
r.clear()  # 清空缓存中的数据
r.set_before(before)  # 设置在数据前面补充的列
r.set_after(after)  # 设置在数据后面补充的列
r.set_head(head)  # 设置表头。只有 csv 和 xlsx 格式支持设置表头
```

### Tips

- add_data() 可以接收 str、int、float、list、tuple、dict 等类型数据
- add_data() 也可以接收这些类型组成的列表，一次插入多条数据
- 在程序退出或崩溃时会自动记录缓存中的数据
- 如果记录器对象是全局变量，退出时不能自动保存，此时会把剩余数据打印出来让程序员自行处理
- 进行采集时，经常除了插入当前采集的数据，还要在这些数据前面或后面插入固定的数据列，可以用 set_before() 和 set_after() 指定这些列，可以接收 str、int、float、list、tuple、dict 等类型数据
- set_before() 和 set_after() 须在 add_data() 前使用，否则添加的数据不会带上这些信息。每次使用这两个方法时都会保存一次数据。
- 如果是新文件且传入的数据是 dict 格式，会自动生成表头
- 指定保存文件的路径不必已经存在，会自动创建
- 使用 set_head() 方法会覆盖第一行数据，对原来没有表头的文件慎用

## Filler 类

Filler 类主要用于对已有数据的表格文件进行填充，也可指定要填写数据的单元格，直接向其填数据。支持 xlsx 和 csv 格式。

使用场景：

- 采集数据时采用分布采集法，先采集 url 存放在文件中，再批量根据 url 采集其中内容，把采集到的内容填到 url 所在行
- 采集数据时采用控制文件方式，用一个文件记录采集的状态、数量，以便断点续爬

### 创建 Filler 对象

```python
f = Filler(path, cache_size, key_cols, begin_row, sign_col, sign, data_col)
# 参数说明：
# path: 保存的文件路径
# cache_size: 每接收多少条记录写入文件，传入0表示不自动保存
# key_cols: 作为关键字的列，可以是多列，从1开始
# begin_row: 数据开始的行，默认表头一行
# sign_col: 用于判断是否已填数据的列，从1开始
# sign: 按这个值判断是否已填数据
# data_col: 要填入数据的第一列，从1开始，不传入时和sign_col一致
```

### Filler 类属性

```python
f.path  # 文件路径
f.cache_size  # 缓存的数据条数
f.data  # 返回当前保存的数据
f.type  # 文件类型
f.key_cols  # 关键字列，可以是多列
f.begin_row  # 数据开始行，默认从第二行开始
f.sign_col  # 用于判断是否已填数据的列，编号从1开始
f.data_col  # 要填入数据的第一列，从1开始，不传入时和sign_col一致
f.keys  # key列内容，第一位为行号，其余为key列的值，eg.[3, '名称', 'id']
f.delimiter  # csv文件分隔符
f.quote_char  # csv文件引用符
f.before  #  补充在前面的列数据
f.after  #  补充在后面的列数据
```

### Filler 类方法

```python
f.set_path()  # 更改文件路径，参数和__init__()一致
f.add_data()  # 插入一条或多条数据，数据第一位为行号或坐标（int或str），第二位开始为数据，数据可以是list, tuple, dict
f.record(new_path)  # 主动保存数据，可指定另存为的路径
f.fill(func, *args)  # 接收一个方法，根据keys自动填充数据。每条key调用一次该方法，并根据方法返回的内容进行填充。方法第一个参数必须是keys，用于接收关键字列
f.set_head(head)  # 设置表头
f.set_link(coord, link, content)  # 为单元格设置超链接
```

### Tips

- add_data() 要插入的数据格式为 [行号或坐标, 数据1, 数据2, ...]
- 上一条数据第一位如果传入 int，可指定行号，如传入 str，即指定坐标。坐标格式：'B3' 或 '3,2'
- 其余 tips 与 Recorder 一致

**注意：** func 返回的数据第一位必须是行号或坐标。

## MapGun 类

MapGun 类用于对一个区域一次写入一整片二维数据。支持 xlsx 和 csv 格式。

### 创建 MapGun 对象

```python
m = MapGun(path, coord, float_coord)  # 传入文件路径，坐标，坐标是否浮动
```

坐标浮动表示填完一批数据后，坐标会移动到数据底部的下一行。

### MapGun 类属性

```python
m.path  # 文件路径
m.cache_size  # 缓存的数据条数，但固定为1
m.coord  # 坐标，形式可以是：'b3'，'3,2'，(3, 2)，[3, 2]
m.float_coord  # 坐标是否随数据增加变化，布尔值
m.type  # 文件类型
m.delimiter  # csv文件分隔符
m.quote_char  # csv文件引用符
m.before  # 补充在前面的列数据
m.after  # 补充在后面的列数据
```

### MapGun 类方法

```python
m.add_data(data, coord)  # 插入二维数据，可同时指定左上角坐标
m.set_before(before)  # 设置在数据前面补充的列
m.set_after(after)  # 设置在数据后面补充的列
m.set_head(head)  # 设置表头
```

### Tips

- MapGun 类的 cache_size 固定为 1，不能修改

## 实用方法

### align_csv()

此方法可补全 csv 文件，使其每行列数一样多，避免 pandas 读取时出错。

```python
from DataRecorder.tootls import align_csv
align_csv(path, encoding, delimiter, quotechar)  # 传入要处理的文件路径及编码，分隔符j
```

