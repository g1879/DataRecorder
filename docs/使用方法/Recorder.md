`Recorder`用途最为广泛，它的功能简单直观高效实用，只做一个动作，就是不断接收数据，按顺序往文件里添加。可以接收单行数据，或二维数据一次写入多行。

它主要支持 csv、xlsx、json、txt 四种格式文件，当目标文件不是这 4 种之一时，按 txt 的方式记录。

# 创建`Recorder`对象

```python
from DataRecorder imoprt Recorder

r = Recorder(path='data.csv', cache_size=500)
```

- `path`：文件路径，若不存在会自动创建，如存在则在最后追加数据。

- `cache_size`：缓存数据条数，到达条数就会写入文件。

# 添加数据

`Recorder`支持非常灵活的数据输入，可一次接收单条或多条数据。

- 接收`str`、`int`、`float`、openpyxl 的`Cell`对象，或其它类型对象时，将其作为一个数据独占一行来记录。

- 接收`list`、`tuple`、`dict`时，会将其中数据记录为一行。

- 接收二维`list`或`tuple`时，会记录为多行。

## 记录单个数据

```python
r.add_data('abc')  # 接收str数据
r.add_data(123)  # 接收int数据
r.add_data(Cell)  # 接收openpyxl Cell对象
```

?>**Tips：**<br>接收 openpyxl 的`Cell`对象时，会记录其`value`值。

## 接收一行多个数据

```python
r.add_data([123, 'abc'])  # 接收list数据
r.add_data((123, 'abc'))  # 接收tuple数据
r.add_data({'a': 1, 'b': 2})  # 接收dict数据
```

?>**Tips：**<br>接收`dict`数据时，会记录其`value`值。如果是一个新文件，会自动根据其`key`值创建表头。

## 接收多行数据

```python
r.add_data([(1, 2), (3, 4)])  # 接收二维list
r.add_data(((1, 2), (3, 4)))  # 接收二维tuple
r.add_data([{'a': 1, 'b': 2}, {'a': 3, 'b': 4}])  # 接收多个dict
```

其中，接收多个`dict`的数据时，如果是新建文件，会自动以`'a'`、`'b'`生成表头。

# 写入文件

## 自动写入

当接收数据到达指定条数，或程序结束时，会触发写入文件动作，同时清空缓存。

如果写入时文件被打开而无法写入，会显式提示，并等待文件关闭再写入。

如果因其它问题导致写入失败，会返回收集到的数据，并打印到控制台。

```python
from DataRecorder import Recorder

r = Recorder('data.csv')
r.add_data(123)  # 程序结束时自动写入文件
```

## 手动写入

可在程序中调用`record()`来提前执行写入动作。

```python
from DataRecorder import Recorder

r = Recorder('data.csv')
r.add_data(123)
r.record()  # 手动调用写入方法
```

## 注意事项

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
    r.add_data('abc')  # 可在程序结束时自动记录

if __name__ == '__main__':
    main()
```

# 文件格式

`Recorder`支持任意后缀文件，针对 xlsx、csv、json、txt 格式作了优化。如果指定这几种格式以外的文件，会自动以 txt 方式进行记录。

- xlsx：
