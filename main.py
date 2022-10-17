# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2022年10月17日15:24:50
# @Author : lian700
# @Email : 853333212@qq.com
# @File : main.py
# @Version : 1.0
# @Description:青年大学习统计脚本

import pandas as pd
from openpyxl.utils import get_column_letter
from pandas import ExcelWriter
import numpy as np


# 不要求但是可以做的人数
exclude_numbers = 2
# 青年大学习期数
period_name = '22期'
# 接龙导出的excel
sheet_name1 = 0
xlsx_name = r'.\接龙统计.xlsx'
# 班级名单路径,sheet_name2 表示excel中多sheet的名字或索引
sheet_name2 = 0
xlsx_name_update = r'.\班级名单.xlsx'
# 保存结果的路径
xlsx_name_save = r'.\output.xlsx'
# 标志性区别列名，判断条件为空
dif_flag = '学习截图'
# 姓名列名
dif_name = '署名'
# 读取excel数据
df = pd.read_excel(xlsx_name, sheet_name=sheet_name1)
complete_persons = df.loc[pd.isnull(df[dif_flag]) == False, [dif_name]]

df2 = pd.read_excel(xlsx_name_update, sheet_name=sheet_name2)

len_all_numbers = len(df2['学号']) - exclude_numbers
# 创建新列
df2.loc[:, period_name] = 0

complete_persons = np.array(complete_persons).squeeze(1)

for i in range(0, len(df2[period_name])):
    # print(df2['姓名'][i])
    if df2['姓名'][i] in complete_persons:
        df2.loc[i, period_name] = 1
len_complete_persons = len(complete_persons)
df2.loc[0, period_name + '完成人数/总人数'] = str(len(complete_persons)) + '/{}人'.format(len_all_numbers)


def to_excel_auto_column_weight(df: pd.DataFrame, writer: ExcelWriter, sheet_name):
    """DataFrame保存为excel并自动设置列宽"""
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    #  计算表头的字符宽度
    column_widths = (
        df.columns.to_series().apply(lambda x: len(x.encode('utf-8'))).values
    )
    #  计算每列的最大字符宽度
    max_widths = (
        df.astype(str).applymap(lambda x: len(x.encode('utf-8'))).agg(max).values
    )
    # 计算整体最大宽度
    widths = np.max([column_widths, max_widths], axis=0)
    # 设置列宽
    worksheet = writer.sheets[sheet_name]
    for i, width in enumerate(widths, 1):
        # openpyxl引擎设置字符宽度时会缩水0.5左右个字符，所以干脆+2使左右都空出一个字宽。
        worksheet.column_dimensions[get_column_letter(i)].width = width + 2


with pd.ExcelWriter(xlsx_name_save) as writer:
    to_excel_auto_column_weight(df2, writer, sheet_name='青年大学习统计')
