# ⭐️简介

本库是一个基于 python 的工具集，用于记录数据到文件。

使用方便，代码简洁， 是一个可靠、省心且实用的工具。

还支持多线程同时写入文件。

**交流QQ群：** 897838127

**联系邮箱：** g1879@qq.com

# ✨️理念

简单，可靠，省心。

# 📕背景

进行数据采集的时候，常常要保存数据到文件，频繁开关文件会影响效率，而如果等采集结束再写入，会有因异常而丢失数据的风险。

因此写了这些工具，只要把数据扔进去，它们能缓存到一定数量再一次写入，减少文件开关次数，而且在程序崩溃或退出时能自动保存，保证数据的可靠。

它使用非常方便，无论何时何地，无论什么格式，只要使用`add_data()`方法把数据存进去即可，语法极其简明扼要，使程序员能更专注业务逻辑。

工具还对表格文件（xlsx、csv）做了很多优化，封装了实用功能，可以使用表格文件方便地实现断点续爬、批量转移数据、指定坐标填写数据等。

# 🍀特性

- 可以缓存数据到一定数量再一次写入，减少文件读写次数，降低开销。
- 支持多线程同时写入数据。
- 可以在程序崩溃时自动保存，失败时显示剩余数据。
- 写入时如文件打开，会自动等待文件关闭再写入，避免数据丢失。
- 对断点续爬提供良好支持。
- 可方便地批量转移数据。
- 可根据字典数据自动创建表头。
- 自动创建文件和路径，减少代码量。

# 🎇概览

这里简要介绍各种工具用途，详细用法请查看使用方法章节。

## 记录器`Recorder`

`Recorder`是最常用的工具，它的功能简单直观高效实用，只做一个动作，就是不断接收数据，按顺序往文件里添加。可以接收单行数据，或二维数据一次写入多行。

它主要支持 csv、xlsx、json、txt 四种格式文件，当目标文件不是这 4 种之一时，按 txt 的方式记录。

```python
from DataRecorder import Recorder

data = ((1, 2, 3, 4), 
        (5, 6, 7, 8))

r = Recorder('data.csv')
r.add_data(data)  # 一次记录多行数据
r.add_data('abc')  # 记录单行数据
```

## 表格填充器`Filler`

`Filler`用于对表格文件填写数据，它每次只接收一行数据，但可以指定填其坐标。它的使用非常灵活，还封装了记录数据处理进度的功能（比如断点续爬）。除此以外，它还能给单元格设置链接。

它只支持 csv 和 xlsx 格式文件。

```python
from DataRecorder import Filler

f = Filler('results.csv')
f.add_data((1, 2, 3, 4), 'a2')  # 从A2单元格开始，写入一行数据
```

## 二维数据记录器`MapGun`

`MapGun`用于对表格文件一次写入一整片二维数据。用法是指定一个坐标，然后以这个坐标为左上角，一次性填入整片数据。

它只支持 csv 和 xlsx 格式文件。

```python
from DataRecorder import MapGun

m = MapGun('results.csv')
data = ((1, 2),
        (3, 4))
m.add_data(data, 'c4')  # 把二维数据填入以c4为左上角的区域中
```

## 二进制数据记录器`ByteRecorder`

`ByteRecorder`用法最简单，它和`Recorder`类似，记录多个数据然后按顺序写入文件。不一样的是它只接收二进制数据，每次`add_data()`只能传入一条数据，没有行的概念。

可以用来和作者的另一个工具 [FlowViewer](https://gitee.com/g1879/FlowViewer) 配合使用，用来获取浏览器加载的文件，或用来记录下载的文件。可指定每个数据写入文件中的位置，以支持多线程下载文件。 

它支持任意文件格式。

```python
from DataRecorder import ByteRecorder

b = ByteRecorder('data.file')
b.add_data(b'xxxxxxxxxxx')  # 向文件写入二进制数据
```

# ☕ 请我喝咖啡

如果本项目对您有所帮助，不妨请作者我喝杯咖啡 ：）

![](https://gitee.com/g1879/DrissionPage-demos/raw/master/pics/code.jpg)