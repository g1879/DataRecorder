## 🎇 概览

本库包含 4 个工具，这里简要介绍它们的用途，详细用法请查看使用方法章节。

各个工具有着相同的使用逻辑：创建对象》添加数据》记录数据。无论何时何地，只要往对象中添加数据即可，其余工作它会自动完成。 并且都支持多线程。

由于逻辑简单，可灵活地应用于多种场景。

```python
from DataRecorder import Recorder

r = Recorder('data.csv')  # 创建对象
r.add_data('abc')  # 添加数据
r.record()  # 记录数据
```

---

### ⚡ 记录器`Recorder`

`Recorder`的功能简单直观高效实用，只做一个动作，就是不断接收数据，按顺序往文件里添加。可以接收单行数据，或二维数据一次写入多行。

支持 csv、xlsx、json、txt 四种格式文件。

```python
from DataRecorder import Recorder

r = Recorder('data.csv')
r.add_data(((1, 2), (3, 4)))  # 用二维数据一次添加多行
r.add_data('abc')  # 添加单行数据
r.record()  # 记录数据
```

---

### ⚡ 表格填充器`Filler`

`Filler`用于对表格文件填写数据，可以指定填其坐标。使用非常灵活，可以指定坐标为左上角，填入一片二维数据。还封装了记录数据处理进度的功能（比如断点续爬）。除此以外，它还能给单元格设置链接。

支持 csv 和 xlsx 格式文件。

```python
from DataRecorder import Filler

f = Filler('data.csv')
f.add_data((1, 2, 3, 4), 'a2')  # 从A2单元格开始，写入一行数据
f.add_data(((1, 2), (3, 4)), 'd4')  # 以D4单元格为左上角，写入一片二维数据
f.record()
```

---

### ⚡ 二进制数据记录器`ByteRecorder`

`ByteRecorder`用法最简单，它和`Recorder`类似，记录多个数据然后按顺序写入文件。不一样的是它只接收二进制数据，每次`add_data()`
只能传入一条数据，没有行的概念。

可指定每个数据写入文件中的位置，以支持多线程下载文件。

支持任意文件格式。

```python
from DataRecorder import ByteRecorder

b = ByteRecorder('data.file')
b.add_data(b'xxxxxxxxxxx')  # 向文件写入二进制数据
b.record()
```

---

### ⚡ 数据库记录器`DBRecorder`

用于向 sqlite 写入数据，用法和`Recorder`一致，支持自动创建数据库、数据表、数据列。

支持 db 格式文件。

```python
from DataRecorder import DBRecorder

d = DBRecorder('data.db')
d.add_data({'name': '张三', 'age': 25}, table='user')  # 插入数据到user表
d.record()
```