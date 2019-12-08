from stats import Stats

def rens_list(cur, lei_xing):
    '''
    从数据库中获取前线人员和后线人员的名单
    并初始化人员信息列表对象
    '''

    year_sql = "SELECT MAX ([年份]) \
                FROM [销售人员业务跟踪表]"

    cur.execute(year_sql)
    max_year = cur.fetchone()[0]

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

        info.set_cai(nian_fen=f"{max_year}")
        info.set_che(nian_fen=f"{max_year}")
        info.set_ren(nian_fen=f"{max_year}")

        rens.append(info)

    return rens