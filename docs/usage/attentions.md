## ❗ 关闭进程的时机

在调试程序的时候经常会手动关闭进程，在关闭的时候务必注意文件写入状态，如果程序正在写入文件，关闭进程可能导致文件损坏，损失之前采集的数据。

当`show_msg`属性为`True`时，程序在开始写入和写入结束的时候，会打印提示语句，请务必注意。

## ❗ 对象的声明

本库的几个工具能够在程序关闭时自动记录缓存中的数据，但以下情况例外：

- 记录器为全局对象，且使用 xlsx 文件

- 记录器为全局对象，且使用 db 文件

- 记录器为全局对象，且接收多线程写入数据

这两种情况下自动记录都会出错。解决方法有以下两种：

- 把记录器对象声明放在一个方法体里

- 显式调用`record()`方法进行写入数据

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
    r.add_data('abc')  # 可在程序结束时自动记录

if __name__ == '__main__':
    main()
```

!!! tip "Tips"
推荐都放在方法体内进行对象声明，避免使用全局对象，养成良好的编码习惯。

## ❗ 慎用负数列号

`Filler`坐标定位方式相当灵活，行号和列号均支持负数，表示从后往前数。

但使用 xlsx 文件时，如果写入的数据长度超过了最大列号（比如在最后一列添加一个长度为 2 的数据），下次写入的时候最大列号就会与上次不一样，导致错位。

```python
from DataRecorder import Filler

f = Filler('test.xlsx', 1, data_col=1)
f.add_data(((1, 2), (3, 4)))  # 写入两列基础数据
f.add_data(('a', 'b'), ('newline', -1))  # 向第二行倒数第一列单元格写入一个数据
f.add_data(('a', 'b'), ('newline', -1))  # 向第二行倒数第一列单元格写入一个数据
f.add_data(('a', 'b'), ('newline', -1))  # 向第二行倒数第一列单元格写入一个数据
f.record()
```

文件内容：

|     |     |     |     |     |
| --- | --- | --- | --- | --- |
| 1   | 2   |     |     |     |
| 3   | 4   |     |     |     |
|     | a   | b   |     |     |
|     |     | a   | b   |     |
|     |     |     | a   | b   |

可见后面 3 条语句虽然是一样的，但每次保存文件就往后移一位。

解决方法有几种：

- 不使用负数列号

- 避免数据长度超出最大列数

- 使用 csv 文件。csv 文件最大列数根据第一行列数确定，下面的行添加列数不影响

- 不自动保存，而使用手动保存，但下次打开文件时依然和这次列数不一样

## ❗ 二维数组的判断

`Recorder`和`Filler`的`add_data()`方法都可以接收二维数组，但也有一维数组里面存在多种数据类型的情况。为了省事，程序以第一个数据的类型来判断接收到的数据是一维还是二维数据。

一维数据把所有数据填在同一行，二维数据每个数据填写一行。

```python
from DataRecorder import Recorder

r = Recorder('test.csv')
data1 = [123, 'abc', (1, 2)]  # 123为单个数据，判断为一维数组
data2 = [(1, 2), 123, 'abc']  # (1, 2)为多个数据，判断为二维数组
r.add_data(data1)
r.add_data(data2)
```

文件内容：

| 123 | abc | (1, 2) |
|:---:|:---:|:------:|
| 1   | 2   |        |
| 123 |     |        |
| abc |     |        |

可见第一个`add_data()`写入了一行数据，第二个写入了 3 行数据。
