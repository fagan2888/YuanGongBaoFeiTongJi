import sqlite3

from openpyxl import Workbook
from openpyxl.styles import (Alignment, Border, Color, Font, NamedStyle,
                             PatternFill, Side)

from test import write, write_zhong_zhi
from rens import rens_list


conn = sqlite3.connect('data.db')
cur = conn.cursor()

wb = Workbook()

title_style = NamedStyle(name='title_style')
title_style.font = Font(name='微软雅黑', size=12, bold=True)
title_style.alignment = Alignment(horizontal='center',wrapText=True)


num_style = NamedStyle(name='num_style')
num_style.font = Font(name='微软雅黑', size=11)
num_style.alignment = Alignment(horizontal='center')
num_style.number_format = '0.00'

str_style = NamedStyle(name='str_style')
str_style.font = Font(name='微软雅黑', size=11)
str_style.alignment = Alignment(horizontal='center')


wb.add_named_style(title_style)
wb.add_named_style(num_style)
wb.add_named_style(str_style)

rens = rens_list(cur, '前线')

# write(wb, rens, '整体')
write(wb, rens, '车险')
write(wb, rens, '非车险')

# write_zhong_zhi(rens, '分公司营业一部')
# write_zhong_zhi(rens, '曲靖')
# write_zhong_zhi(rens, '文山')
# write_zhong_zhi(rens, '大理')
# write_zhong_zhi(rens, '保山')
# write_zhong_zhi(rens, '版纳')
# write_zhong_zhi(rens, '怒江')
# write_zhong_zhi(rens, '昭通')

rens = rens_list(cur, '后线')
write(wb, rens, '整体', '后线')
# wb.move_sheet('后线人员整体保费排名', offset=-8)

wb.remove(wb['Sheet'])
wb.save("2019年销售人员业务跟踪表.xlsx")

cur.close()
conn.close()