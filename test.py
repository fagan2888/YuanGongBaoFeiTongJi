import sqlite3
from openpyxl import Workbook
from stats import Stats


def rens_list(lei_xing):
    if lei_xing == '前线':
        str_sql = "SELECT 中支公司, 营销服务部, 业务员, 入司时间, 职级 FROM 前线"
    else:
        str_sql = "SELECT 中支公司, 营销服务部, 业务员, 入司时间 FROM 后线"

    cur.execute(str_sql)

    rens = []

    for values in cur.fetchall():
        info = Stats(cur)
        info.zhong_zhi = values[0][7:]
        info.ji_gou = values[1][11:]
        info.name = values[2][10:]
        info.ru_si_shi_jian = values[3][:10]
        if lei_xing == '前线':
            info.zhi_ji = values[4][16:]
        info.set_cai(nian_fen='2019')
        info.set_che(nian_fen='2019')
        info.set_ren(nian_fen='2019')

        rens.append(info)

    return rens


def write(rens, xian_zhong, lei_xing='前线'):
    if xian_zhong == '整体':
        rens_sort = sorted(rens, key=lambda ren: ren.zheng_ti, reverse=True)
    elif xian_zhong == '车险':
        rens_sort = sorted(rens, key=lambda ren: ren.che, reverse=True)
    else:
        rens_sort = sorted(rens, key=lambda ren: ren.fei_che, reverse=True)

    ws = wb.create_sheet(title=f'{lei_xing}人员{xian_zhong}保费排名')

    row = [f'{lei_xing}人员{xian_zhong}保费排名']
    ws.append(row)

    if lei_xing == '前线':
        row = ['排名',
               '姓名',
               f'{xian_zhong}保费（万元）',
               '机构',
               '职级',
               '入司时间']
    else:
        row = ['排名',
               '姓名',
               f'{xian_zhong}保费（万元）',
               '机构',
               '入司时间']
    ws.append(row)

    i = 1

    for ren in rens_sort:
        row = [i, ren.name]
        if xian_zhong == '整体':
            row.append(ren.zheng_ti)
        elif xian_zhong == '车险':
            row.append(ren.che)
        else:
            row.append(ren.fei_che)
        row.append(ren.ji_gou)
        if lei_xing == '前线':
            row.append(ren.zhi_ji)
        row.append(ren.ru_si_shi_jian)
        ws.append(row)
        i += 1


def write_zhong_zhi(rens, zhong_zhi):
    new_rens = []
    for ren in rens:
        if zhong_zhi in ren.zhong_zhi:
            new_rens.append(ren)

    rens_sort = sorted(new_rens, key=lambda ren: ren.zheng_ti, reverse=True)

    ws = wb.create_sheet(title=f'{zhong_zhi}前线人员整体保费跟踪表')

    row = [f'{zhong_zhi}前线人员整体保费跟踪表']
    ws.append(row)

    row = ['排名',
           '姓名',
           '保费（万元）',
           '机构',
           '职级',
           '入司时间']

    ws.append(row)

    i = 1

    for ren in rens_sort:
        row = [i,
               ren.name,
               ren.zheng_ti,
               ren.ji_gou,
               ren.zhi_ji,
               ren.ru_si_shi_jian,
               ren.bao_fei(xian_zhong='整体', nian_fen='2019', yue_fen='01')]
        ws.append(row)
        i += 1


conn = sqlite3.connect('data.db')
cur = conn.cursor()

wb = Workbook()

rens = rens_list('前线')

write(rens, '整体')
write(rens, '车险')
write(rens, '非车险')

write_zhong_zhi(rens, '曲靖')

# rens = rens_list('后线')
# write(rens, '整体', '后线')
# wb.move_sheet('后线人员整体保费排名', offset=-1)

wb.remove(wb['Sheet'])
wb.save("2019年销售人员业务跟踪表.xlsx")

cur.close()
conn.close()
