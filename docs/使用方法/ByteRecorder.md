`ByteRecorder`专门用于记录二进制数据到文件。

用法和`Recorder`类似，每次`add_data`只接收一段二进制数据，可以指定数据在文件中要写入的位置。

`ByteRecorder`可用于多线程文件下载、记录监听数据流等场景。支持任意后缀文件。

**优点：**

- 高度集成功能，简化代码逻辑

- 接收和写入数据分离，避免多线程时写入文件冲突

# 🐘创建对象

```python
from DataRecorder import ByteRecorder

b = ByteRecorder(path='data.file', cache_size=50)
```

初始化参数：

- `path`：文件路径，若不存在会自动创建。

- `cache_size`：缓存数据条数，到达条数自动写入文件。为 0 时不自动写入。

# ➕添加数据

`ByteRecorder`只能接收二进制格式数据，且一次接收一条数据，可指定数据在文件中的位置。如不指定，默认添加到文件末尾。

```python
b.add_data(b'abcd')  # 将一条数据添加到文件末尾
b.add_data(b'1234', seek=2)  # 在第2个字节的位置添加一条数据，0开始
```

!>**注意：**<br>指定位置添加数据，会覆盖该位置后面已有数据。

文件内容：

```
ab1234
```

可见该文件首先添加了数据 abcd，然后在 2 的位置添加了 1234，覆盖了后面的 cd。

# 💾写入文件

这部分内容与`Recorder`一致。同样支持自动写入和手动写入。

## ⚠️注意事项

- 如果对象为全局对象，那么在使用 xlsx 或多线程写入的时候，退出自动记录功能会报错，须显式调用`record()`方法，或把记录器声明放在一个方法内。

- 切勿在文件正在写入的时候关闭进程，以免造成文件损坏。

详见[《注意事项》](%E6%B3%A8%E6%84%8F%E4%BA%8B%E9%A1%B9.md)一节。

# 📌示例

这段示例演示用多线程写入多行数据，来模拟多线程下载文件。

```python
from DataRecorder import ByteRecorder
from threading import Thread

def add_data(num: int, recorder: ByteRecorder):
    data = str(num) * 10 + '\n'
    data = bytes(data, encoding='utf-8')
    seek = num * 11

    recorder.add_data(data, seek)

def main():
    b = ByteRecorder('data.csv')
    for i in range(5):
        Thread(target=add_data, args=(i, b)).start()

if __name__ == '__main__':
    main()
```

!>**注意：**<br>多线程写入数据时，ByteRecorder 对象的声明要放在方法体内。

文件内容：

```
0000000000
1111111111
2222222222
3333333333
4444444444
```

# 🔣`ByteRecorder`对象的属性

## `path`

此属性以字符串方式返回当前记录的文件路径。可赋值设置。更改时会自动保存缓存数据到文件。

## `type`

此属性以字符串方式返回当前文件的类型。可赋值设置，即可无视文件后缀指定记录格式。

## `data`

此属性以`list`方式返回当前保存在缓存里的数据。

## `cache_size`

此属性返回缓存的大小，表示记录的条数。可赋值设置。

## `show_msg`

此属性用于设置是否打印程序运行时产生的提示信息。

# ♾️`ByteRecorder`对象的方法

## `add_data()`

此方法用于添加数据到缓存，只能接收`bytes`类型数据，可指定数据在文件中的位置。

参数：

- `data`：`bytes`类型数据
- `seek`：在文件中的位置，`None`表示最后

返回：`None`

## `record()`

此方法用于把数据记录到文件，然后清空缓存。可把数据保存到一个新文件。

参数：

- `new_path`：保存到新文件的路径

返回：成功时以文本方式返回文件路径，失败时返回未保存的数据

## `set_path()`

此方法用于设置文件路径，更改文件路径时会自动保存已有缓存。

参数：

- `path`：文件路径，可以是`str`或`Path`对象
- `file_type`：要设置的文件类型，为空则从文件名中获取

返回：`None`

## `clear()`

此方法用于清空现有缓存。

参数：无

返回：`None`
