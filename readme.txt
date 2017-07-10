使用前修改配置文件config.ini,配置文件格式不要改动，只用修改等号后面的值即可，每个字段代表的意思分别是：

username  登陆时的用户名；
password  登陆时的密码；
subject_name 查询用的项目名 注：不要有空格；
product_name 查询用的产品名 注：不要有空格
given_date 新增bug的判断日期；
circle 统计周期，即多长时间统计一次；
delay_keyword 定义延期bug的关键字
filter_term 筛选条件
classify_keywords 按处理状态分类时所需要查找的关键字
kind_keywords 按处理状态分类时的类别
classify_switch 是否使用按处理状态分类,True为使用，False为不使用，其他值为无效。
delay_switch 是否输出延期bug列表,True为输出，False为不输出，其他值无效。

其中subject_name和product_name只能选填一个，另一个要填成null。不能两个都填。

其中classify_keywords和kind_keywords需要按样例进行填写：用中括号括住，每个关键字用双引号括住，
两个关键字之间使用逗号隔开。分类关键字和类别关键字的位置要是对应的，比如标题里有“验证”两个字的，
对应的类型为“待验证”，这两个关键字的位置要对应。


使用完毕后，将password的值改为*, 避免泄漏

使用时运行measure.py即可

搜索结果.xls文件是数据来源，不要动

项目名称.xlsx文件是最后结果文件，项目名称.txt文件是统计结果记录文件，也不用动。

程序运行时可以做其他工作，不要对程序运行窗口和firefox浏览器进行操作即可。

本程序暂时还有一些问题，但是影响不大，如有其他问题欢迎反馈到yinguangchao@kedacom.com