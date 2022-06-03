# ⭐️简介

本库是一个基于 python 的工具集，用于记录数据到文件。

使用方便，代码简洁， 是一个非常实用的工具。还支持多线程同时写入文件。

## ✨️理念

简单，可靠，人性化

## 📕背景

进行数据采集的时候，常常要保存数据到文件，频繁开关文件会影响效率，而如果等采集结束再写入，会有因异常而丢失数据的风险。

因此写了一个工具，只要把数据扔进去，它能缓存到一定数量再一次写入，减少文件开关次数，而且在程序崩溃或退出时能自动保存，保证数据的可靠。

它使用非常方便，无论何时何地，只要使用`add_data()`方法把数据存进去即可，语法极其简明扼要，使程序员能更专注业务逻辑。

工具还对表格文件（xlsx、csv）做了很多优化，封装了实用功能，可以使用表格文件方便地实现断点续爬、批量转移数据、指定坐标填写数据等。

## 🍀特性

- 可以缓存数据到一定数量再一次写入，减少文件读写次数，降低开销。
- 支持多线程同时写入数据。
- 可以在程序崩溃时自动保存，失败时显示剩余数据。
- 写入时如文件打开，会自动等待文件关闭再写入，避免数据丢失。
- 对断点续爬提供良好支持。
- 可方便地批量转移数据。
- 自动创建文件和路径，减少代码量。

# 🎇简单演示

下面这些示例可直接运行（除了第3个）

1、填入二维数组，记录多行数据

```python
from DataRecorder import Recorder

r = Recorder('results.csv')  # 声明记录器（会自动创建该文件）
data = [(1, 2, 3, 4) for _ in range(100)]  # 产生一个100行的二维数组
r.add_data(data)  # 把数据记录到文件
```

2、对表格坐标录入数据

```python
from DataRecorder import Filler

f = Filler('results.csv')
f.add_data((1, 2, 3, 4), 'a2')  # 向 A2 单元格写入数据
f.add_data((1, 2, 3, 4), (3, 4))  # 向第3行第4列开始的单元格写入该数据
f.add_data((1, 2, 3, 4), 'c')  # 新增一行并在该行C列开始记录该数据
f.add_data((1, 2, 3, 4), (-3, 4))  # 在倒数第3行第4列开始记录该数据
```

3、根据某列数据对其它列填充数据

```python
from DataRecorder import Filler

f = Filler('results.csv', key_cols='A', sign_col='B')
# =============方法一=============
for key in f.keys:  # 所有未填充行的key列
    data = do_sth(key, *args)  # 处理数据的方法，第一个参数必须是接收key值
    f.add_data(data, key)

# =============方法二=============
f.fill(do_sth, *args)  # 调用处理数据方法，自动填充数据
```

4、向表格文件写入一整片二维数据

```python
from DataRecorder import MapGun

m = MapGun('results.csv')
data = ((1, 2),
        (3, 4))
m.add_data(data, 'c4')  # 把二维数据填入以 c4 为左上角的区域中
```

5、多线程同时写入

```python
from DataRecorder import Recorder
from threading import Thread

def add(recorder: Recorder):
    recorder.add_data((1, 2, 3, 4))

def main():
    r = Recorder('results.csv')
    for _ in range(5):  # 创建 5 个线程，每个都向文件写入数据
        Thread(target=add, args=(r,)).start()

if __name__ == '__main__':
    main()
```

# 注意事项

如果记录文件格式为 xlsx，程序退出自动记录功能无法使用，须显式调用`record()`方法，或把记录器声明放在一个方法内。

错误做法：

```python
from DataRecorder import Recorder

r = Recorder('test.xlsx')  # 使用 xlsx 格式
r.add_data('abc')
# 不显式调用 record()方法
```

正确做法1，显式调用`record()`方法：

```python
from DataRecorder import Recorder

r = Recorder('test.xlsx')  # 使用 xlsx 格式
r.add_data('abc')
r.record()  # 显式调用 record()方法
```

正确做法2，把对象声明放在方法体内：

```python
from DataRecorder import Recorder

def main():
    r = Recorder('test.xlsx')  # r 的声明放在方法体内
    r.add_data('abc')

if __name__ == '__main__':
    main()
```

# 安装与导入

## 安装

```
pip install DataRecorder
```

## 导入

```python
from DataRecorder import Recorder  # 记录器
from DataRecorder import Filler  # 填充器
from DataRecorder import MapGun  # 范围填充器
```

# 使用方法

## 记录器：Recorder 类

Recorder 用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销。退出时能自动记录数据，避免因异常丢失。支持 xlsx、csv、json、txt 格式。如果指定这几种格式以外的文件，会自动以 txt 方式进行记录。

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

## 填充器：Filler 类

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

## 范围填充器：MapGun 类

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

# ☕ 请我喝咖啡

如果本项目对您有所帮助，不妨请作者我喝杯咖啡 ：）

![](https://gitee.com/g1879/DrissionPage-demos/raw/master/pics/code.jpg)
