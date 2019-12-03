import sqlite3


class Infos():

    def __init__(self):
        self._zhong_zhi = None
        self._ji_gou = None
        self._name = None
        self._ru_si_shi_jian = None
        self._che_bao_fei = None

    @property
    def zhong_zhi(self):
        return self._zhong_zhi

    @zhong_zhi.setter
    def zhong_zhi(self, str):
        self._zhong_zhi = str

    @property
    def ji_gou(self):
        return self._ji_gou

    @ji_gou.setter
    def ji_gou(self, str):
        self._ji_gou = str

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, str):
        self._name = str

    @property
    def ru_si_shi_jian(self):
        return self._ru_si_shi_jian

    @ru_si_shi_jian.setter
    def ru_si_shi_jian(self, str):
        self._ru_si_shi_jian = str

    @property
    def che_bao_fei(self):
        return self._che_bao_fei

    @che_bao_fei.setter
    def che_bao_fei(self, value):
        self._che_bao_fei = value

    @property
    def cai_bao_fei(self):
        return self._cai_bao_fei

    @cai_bao_fei.setter
    def cai_bao_fei(self, value):
        self._cai_bao_fei = value

    @property
    def ren_bao_fei(self):
        return self._ren_bao_fei

    @ren_bao_fei.setter
    def ren_bao_fei(self, value):
        self._ren_bao_fei = value


conn = sqlite3.connect('data.db')
cur = conn.cursor()

str_sql = "SELECT 中支公司, 营销服务部, 业务员, 入司时间 FROM 前线"

cur.execute(str_sql)

rens = []

for values in cur.fetchall():
    info = Infos()
    info.zhong_zhi = values[0][7:]
    info.ji_gou = values[1][11:]
    info.name = values[2][10:]
    info.ru_si_shi_jian = values[3][:10]

    rens.append(info)

i = 0

while i < len(rens):
    str_sql = "SELECT SUM([签单保费/批改保费]) FROM 销售人员业务跟踪表 "

    str_sql += f"WHERE 业务员 like '%{rens[i].name}%' \
                 AND 中心支公司 like '%{rens[i].zhong_zhi}' \
                 AND [车险/财产险/人身险] = '车险'"

    cur.execute(str_sql)
    rens[i].che_bao_fei = cur.fetchone()[0]
    if rens[i].che_bao_fei is None:
        rens[i].che_bao_fei = 0
    else:
        rens[i].che_bao_fei /= 10000

    i += 1

i = 0
while i < len(rens):
    print(rens[i].zhong_zhi,
          rens[i].ji_gou,
          rens[i].name,
          rens[i].ru_si_shi_jian,
          rens[i].che_bao_fei)

    i += 1
