本库是一个基于 python 的工具集，用于记录数据到文件。

使用方便，代码简洁， 是一个可靠、省心且实用的工具。

支持多线程同时写入文件。

**交流QQ群：** 897838127[已满]、558778073

**联系邮箱：** g1879@qq.com

# ✨️理念

简单，可靠，省心。

# 📕背景

进行数据采集的时候，常常要保存数据到文件，频繁开关文件会影响效率，而如果等采集结束再写入，会有因异常而丢失数据的风险。

因此写了这些工具，只要把数据扔进去，它们能缓存到一定数量再一次写入，减少文件开关次数，而且在程序崩溃或退出时能自动保存，保证数据的可靠。

它们使用非常方便，无论何时何地，无论什么格式，只要使用`add_data()`方法把数据存进去即可，语法极其简明扼要，使程序员能更专注业务逻辑。

它们还相当可靠，作者曾一次过连续记录超过 300 万条数据，也曾 50 个线程同时运行写入数万条数据到一个文件，依然轻松胜任。

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

# ☕ 请我喝咖啡

如果本项目对您有所帮助，不妨请作者我喝杯咖啡 ：）

![](http://g1879.gitee.io/drissionpagedocs/imgs/code.jpg)