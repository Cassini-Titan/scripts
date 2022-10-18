平时工作用到的脚本：

## 日常

-   `surprise.py`
    -   用途：抽奖，或者抽人参加某些强制的活动
    -   特性：
        1. 使用`csv`保存记录以保证更高的兼容性
        2. 支持从`Sheet1`的表头为`Record`的`A`列中读取人名并更新记录的`csv`文件
        3. 算法使用`Count`和`LastTime`完成优先队列，对于两者都相同的列表进行`shuffle`以保证随机性

## 学习

-   `nbconvert.py`
    -   用途：转换`jupyter notebook`文件为`python`文件

## 工具

-   `cmder.py`
    -   用途：提供各种颜色的输出与执行命令的函数
