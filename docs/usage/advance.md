## 📰 设置表头

用表格文件做数据采集的时候，要设置表头，我们可以把表头作为第一条数据通过`add_data()`方法添加到文件。但我们通常会使用同一段代码对一个文件进行多批次写入，如果每次都执行到这个语句，就会写入多个表头行。

因此设计了一个设置表头的方法`set.head()`，`Recorder`还有一个自动设置表头的功能。

### 📜 自动设置

自动设置只有`Recorder`支持，同时满足两个条件就能自动激活：

- 指定不存在的 csv 或 xlsx 文件

- 数据格式为`dict`

写入数据时，程序会自动创建文件，同时会把第一条数据的`key`值作为表头添加到文件。

```python
from DataRecorder import Recorder

r = Recorder('data.csv')
data = {'姓名':'张三', '性别':'男'}
r.add_data(data)
```

文件内容：

| 姓名  | 性别  |
| --- | --- |
| 张三  | 男   |

---

### 📜 手动设置

`set.head()`方法可以为 xlsx 或 csv 文件设置表头，无论文件是否已创建。任何时候都可以使用。

`Recorder`和`Filler`都支持此方法。

```python
from DataRecorder import Recorder

r = Recorder('data.csv')
r.set.head(('姓名', '性别'))
r.add_data(('张三', '男'))
```

!!! warning "注意"
`set.head()`会覆盖第一行内容，使用时要注意。

---

## 🔖 为每条数据添加固定前后缀

有些时候，数据采集的时候要为该批次每条数据都录入同一个内容，比如日期。我们当然可以直接在数据里面加上这些列，但有一个更方便的方法，可以减少这些内容对采集代码的干扰。就是`before`和`after`属性。

顾名思义，`before`是在每条数据记录的时候在前面添加一些固定列，`after`则是在后面。

`Recorder`和`Filler`支持这两个属性，`ByteRecorder`不支持。

下面这个示例，采集 gitee 开源项目前 5 页。此示例使用 [DrissionPage](http://g1879.gitee.io/drissionpage) 进行页面访问和解析。

```python
from DrissionPage import SessionPage
from DataRecorder import Recorder

def get_list(page, recorder):
    """获取一页信息并添加到记录器"""
    p = SessionPage()  # 创建页面对象
    url = f'https://gitee.com/explore/all?page={page}'
    p.get(url)  # 访问页面
    rows = p('.ui relaxed divided items explore-repo__list').eles('.item')
    for row in rows:  # 遍历所有行
        data = {  # 产生一行数据
            'page': page,
            'title': row('.title project-namespace-path').text,
            'content': row('.project-desc mb-1').text,
            'stars': row('.stars-count').text
        }
        recorder.add_data(data)  # 把一条数据放入记录器

def main():
    r = Recorder('data.xlsx')  # 创建记录器
    for i in range(1, 6):  # 遍历5页
        get_list(i, r)

if __name__ == '__main__':
    main()
```

文件内容：

| page | title           | content                                                                                             | stars |
| ---- | --------------- | --------------------------------------------------------------------------------------------------- | ----- |
| 1    | RTduino/RTduino | RT-Thread的Arduino生态兼容层                                                                              | 3     |
| 1    | zu1k/nali       | An offline tool for querying IP geographic information and CDN provider.一个查询IP地理信息和CDN服务提供商的离线终端工具. | 2     |
| 1    | 徐小夕/dooringx    | 开箱即用的H5可视化搭建框架，轻松搭建自己的可视化搭建平台                                                                       | 31    |
|      | 以下省略。。。         |                                                                                                     |       |

我们希望在前面增加一列用于记录采集日期，在后面添加一列记录采集人，这些列每一行都一样，因此可以统一写在`before`和`after`属性中。

```python
r = Recorder('data.xlsx')
r.set.before({'date': '2022-06-10'})
r.set.after({'staff': 'g1879'})
```

现在删掉刚才的文件重新执行采集，得到以下文件内容：

| date       | page | title           | content                                                                                             | stars | staff |
| ---------- | ---- | --------------- | --------------------------------------------------------------------------------------------------- | ----- | ----- |
| 2022-06-10 | 1    | RTduino/RTduino | RT-Thread的Arduino生态兼容层                                                                              | 3     | g1879 |
| 2022-06-10 | 1    | zu1k/nali       | An offline tool for querying IP geographic information and CDN provider.一个查询IP地理信息和CDN服务提供商的离线终端工具. | 2     | g1879 |
| 2022-06-10 | 1    | 徐小夕/dooringx    | 开箱即用的H5可视化搭建框架，轻松搭建自己的可视化搭建平台                                                                       | 31    | g1879 |
|            |      | 以下省略。。。         |                                                                                                     |       |       |

也许您已经注意到，代码没有执行`set.head()`，但前后两列的`key`值也加到表头里去了。

而且这两个属性还支持多列，可以接收`list`、`tuple`、`dict`等，添加前后列。

!!! warning "注意"
    `DBRecorder`设置的`after`和`before`如果不是`dict`格式，而`add_data()`传入的数据是`dict`格式，`after`和`before`设置会被忽略。

---

## 🧶 多线程写入文件

做数据采集的时候经常用到多线程，多线程同时记录到文件容易产生冲突。本库为多线程数据采集提供了支持，只要在主线程中创建记录器对象，分发到各个采集线程中使用即可。用法和单线程一致，可避免写入冲突。

现在继续用上面的示例，创建 5 个线程分别爬取 gitee 开源项目前 5 页，每个线程接收同一个记录器对象，往里面添加数据。只要一个小改动，把对`get_list()`的调用放入线程里即可。

```python
# 单线程：
def main():
    r = Recorder('data.xlsx')  # 创建记录器
    for i in range(1, 6):  # 编历5页
        get_list(i, r)

# 多线程：
def main():
    r = Recorder('data.xlsx')  # 创建记录器
    for i in range(1, 6):  # 创建5个线程分别获取1-5页
        Thread(target=get_list, args=(i, r)).start()
```

文件内容：

| page | title                                    | content                                                                                                | stars |
| ---- | ---------------------------------------- | ------------------------------------------------------------------------------------------------------ | ----- |
| 4    | RichardoMu/yolov5-mask-detect            | A C++ implementation of Yolov5 to detect mask running in Jetson Xavier nx and Jetson nano.             | 17    |
| 4    | RichardoMu/yolov5-smoke-detection-python | A Python implementation of Yolov5 to detect whether peaple smoking in Jetson Xavier nx and Jetson nano | 4     |
| 4    | RichardoMu/yolov5-smoking-detect         | A C++ implementation of Yolov5 to detect peaple smoking running in Jetson Xavier nx and Jetson nano.   | 6     |
|      | 以下省略。。。                                  |                                                                                                        |       |

!!! warning "注意"
因为是多线程采集，文件内容并不是按顺序添加。
