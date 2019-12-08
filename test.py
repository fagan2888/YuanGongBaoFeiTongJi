from stats import Stats


def write(wb, rens, xian_zhong, lei_xing='前线'):
    '''
    将分公司全体人员数据按险种写入到Excel中
    按保费规模进行排名
    '''
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

    if lei_xing == '前线':
        ws.merge_cells("A1:F1")
    else:
        ws.merge_cells("A1:E1")

    ws['A1'].style = 'title_style'

    r = 1
    for row in ws.rows:
        c = 1
        for cell in row:
            if r <=2:
                cell.style = 'title_style'
            else:
                if c == 3:
                    cell.style = 'num_style'
                else:
                    cell.style = 'str_style'
            c += 1
        r +=1

    print(f'{lei_xing}人员{xian_zhong}保费排名表写入完成')




def write_zhong_zhi(rens, zhong_zhi):
    '''
    按中支为单位，将前线人员信息写入到Excel中
    以个人最新一年整体保费规模排名
    '''

    # 筛选出相应中支的人员
    new_rens = []
    for ren in rens:
        if zhong_zhi in ren.zhong_zhi:
            new_rens.append(ren)

    # 以整体保费规模进行排序
    rens_sort = sorted(new_rens, key=lambda ren: ren.zheng_ti, reverse=True)

    if zhong_zhi == '分公司营业一部':
        zhong_zhi = '昆明'
    ws = wb.create_sheet(title=f'{zhong_zhi}')

    row = [f'{zhong_zhi}前线人员整体保费跟踪表']
    ws.append(row)

    row = ['排名',
           '姓名',
           '机构',
           '职级',
           '入司时间']

    year_sql = "SELECT [年份] \
                FROM [销售人员业务跟踪表] \
                GROUP  BY [年份] \
                ORDER  BY [年份] DESC"

    cur.execute(year_sql)
    year = []
    for y_tupe in cur.fetchall():
        for y in y_tupe:
            year.append(y)

    for y in year:
        row.append(f'{y}年\n保费（万元）')

    for y in year:
        month_sql = f"SELECT [月份] \
                     FROM [销售人员业务跟踪表] \
                     WHERE [年份] = '{y}' \
                     GROUP  BY [月份] \
                     ORDER  BY [月份] DESC"

        cur.execute(month_sql)
        for m_tupe in cur.fetchall():
            for m in m_tupe:
                row.append(f"{y}年{m}月\n保费（万元）")

    ws.append(row)

    i = 1

    for ren in rens_sort:
        row = [i,
               ren.name,
               ren.ji_gou,
               ren.zhi_ji,
               ren.ru_si_shi_jian]
        for y in year:
            row.append(ren.bao_fei(nian_fen=y))

        for y in year:
            month_sql = f"SELECT [月份] \
                        FROM [销售人员业务跟踪表] \
                        WHERE [年份] = '{y}' \
                        GROUP  BY [月份] \
                        ORDER  BY [月份] DESC"

            cur.execute(month_sql)
            for m_tupe in cur.fetchall():
                for m in m_tupe:
                    row.append(ren.bao_fei(nian_fen=y, yue_fen=m))

        ws.append(row)
        i += 1

    # title_sytle = NamedStyle(name='title_sytle')
    # title_sytle.font(name='微软雅黑', size=14, bold=True)
