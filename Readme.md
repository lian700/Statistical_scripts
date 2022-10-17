# 使用说明
```
$ pip install pandas
$ pip install openpyxl
```
### 同级目录
* 班级名单替换为自己班级的名单
* 接龙统计使用的是微信接龙管家导出的EXCEL
* 也可以在脚本里自己更改路径名称
* 不使用接龙统计，只要具有标志性判断标志即可，例如
```python
# 脚本通过是否有学习截图来判断是否完成青年大学习
# 标志性区别列名，判断条件为字符串空
dif_flag = '学习截图'
# 姓名列名
dif_name = '署名'
# 不是做空判断的需要修改下面的语句
complete_persons = df.loc[pd.isnull(df[dif_flag]) == False, [dif_name]]
```
### 修改脚本
```python
# 不需要做的人数
exclude_numbers = 1
# 期数
period_name = '22期'
```
### 运行
```
$ python main.py
```
### 后续使用
    第一次用了之后，后面几期23期，24期等需要将 "output.xlsx" 重命名
    为 "班级名单.xlsx" 并覆盖原有的 "班级名单.xlsx"，这是为了保存历史的统计信息。
