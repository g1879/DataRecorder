`DBRecorder`为 3.1 版本新增功能，目前处于测试阶段，只支持 sqlite，文件格式为 db。

它专门用于记录数据到数据库，用法与`Recorder`一致。

可自动创建数据库文件、表、列。

# 🐘创建对象

```python
from DataRecorder imoprt DBRecorder

d = DBRecorder(path='data.db', cache_size=500, table='table1')
```

初始化参数：

- `path`：数据库文件路径，若不存在会自动创建。

- `cache_size`：缓存数据条数，到达条数自动写入文件。为 0 时不自动写入。

- `table`：默认表格名称，可不填。

# ➕添加数据

`DBRecorder`使用方法与`Recorder`基本一致，具体查看该章节，这里展示不一样的地方。

## ✔指定默认表

可以对`DBRecorder`对象设置默认表，每条数据都会添加到该表。

```python
from DataRecorder import DBRecorder

d = DBRecorder(path='data.db', table='user')  # 声明时指定
# 或
d = DBRecorder(path='data.db')
d.table = 'user'  # 对对象赋值
```

## ✔每条数据指定表

也可以在每次添加数据时指定表名。

```python
from  DataRecorder import DBRecorder

d = DBRecorder(path='data.db')
d.add_data({'name':'张三', 'age':20}, table='user1')
d.add_data({'name':'李四', 'age':25}, table='user2')
d.record()
```

## ✔数据格式

与`Recorder`一样，`DBRecorder`也支持灵活的数据格式，但推荐使用`dict`格式。

如上面的示例，`dict`格式可以为不同列指定数据。如果不存在该列，程序会自动创建。

如果使用其它格式，则会从左到右按顺序填入到表的列里。

接收`tuple`格式数据：

```python
from  DataRecorder import DBRecorder

d = DBRecorder(path='data.db', table='user')
d.add_data(('张三', 20))
d.record()
```

## ✔接收多行数据

`DBRecorder`也可以接收二维数据，并逐行添加到指定表中。

```python
d.add_data(({'name':'张三', 'age':20}, 
            {'name':'李四', 'age':25}), table='user')  # 接收多个dict
```

# ⚠️注意事项

- 如果对象为全局对象，那么退出自动记录功能会报错，须显式调用`record()`方法，或把记录器声明放在一个方法内。

- 切勿在文件正在写入的时候关闭进程，以免造成文件损坏。

详见[《注意事项》](注意事项.md)一节。

# 🔣`DBRecorder`对象的属性

## `path`

此属性以字符串方式返回当前记录的文件路径。可赋值设置。更改时会自动保存缓存数据到文件。

## `type`

此属性以字符串方式返回当前文件的类型。可赋值设置，即可无视文件后缀指定记录格式。

## `table`

此属性为数据库默认表名。

## `data`

此属性以`list`方式返回当前保存在缓存里的数据。

## `cache_size`

此属性返回缓存的大小，表示记录的条数。可赋值设置。

## `before`

此属性返回当前对象设置的`before`参数内容，`before`参数内容的用法将在 “进阶用法” 章节说明。

## `after`

此属性返回当前对象设置的`after`参数内容，`after`参数内容的用法将在 “进阶用法” 章节说明。

## `show_msg`

此属性用于设置是否打印程序运行时产生的提示信息。

# ♾️`DBRecorder`对象的方法

## `add_data()`

此方法用于添加数据到缓存，可接收一个、一行或多行数据。

参数：

- `data`：可接收任意格式，接收一维`list`、`tuple`和`dict`时记录为一行，接收二维数据时记录为多行

返回：`None`

## `record()`

此方法用于把数据记录到文件，然后清空缓存。可把数据保存到一个新文件。

参数：

- `new_path`：保存到新文件的路径

返回：成功时以文本方式返回文件路径，失败时返回未保存的数据

## `set_path()`

此方法用于设置文件路径，更改时会自动保存缓存数据到文件。

参数：

- `path`：文件路径，可以是`str`或`Path`对象
- `file_type`：要设置的文件类型，为空则从文件名中获取

返回：`None`

`set_table()`

此方法用于设置数据库默认表名。

参数：

- `table`：表名

返回：`None`

## `clear()`

此方法用于清空现有缓存。

参数：无

返回：`None`

## `set_before()`

此方法用于设置对象`before`参数内容，设置前会先保存缓存数据。

`before`内容的用法将在 “进阶用法” 章节说明。

参数：

- `before`：每行数据前要添加的内容，单个数据或列表数据皆可

返回：`None`

## `set_after()`

此方法用于设置对象`after`参数内容，设置前会先保存缓存数据。

`after`内容的用法将在 “进阶用法” 章节说明。

参数：

- `after`：每行数据后要添加的内容，单个数据或列表数据皆可

返回：`None`
