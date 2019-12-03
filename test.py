import sqlite3
from operator import itemgetter, attrgetter


class Infos():

    def __init__(self):
        self._zhong_zhi = None
        self._ji_gou = None
        self._name = None
        self._ru_si_shi_jian = None
        self._che_bao_fei = None
        self._cai_bao_fei = None
        self._ren_bao_fei = None

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

    def __repr__(self):
        return repr((self.zhong_zhi,
                     self.ji_gou,
                     self.name,
                     self.ru_si_shi_jian,
                     self.che_bao_fei,
                     self.cai_bao_fei,
                     self.ren_bao_fei))

def bao_fei(name, zhong_zhi, xian_zhong):
    str_sql = f"SELECT SUM([签单保费/批改保费]) \
                FROM 销售人员业务跟踪表 \
                WHERE 业务员 like '%{name}%' \
                AND 中心支公司 like '%{zhong_zhi}' \
                AND [车险/财产险/人身险] = '{xian_zhong}'"

    cur.execute(str_sql)
    value = cur.fetchone()[0]
    if value is None:
        value = 0.0
    else:
        value /= 10000

    return value


conn = sqlite3.connect('data.db')
cur = conn.cursor()

str_sql = "SELECT 中支公司, 营销服务部, 业务员, 入司时间 FROM 前线"

cur.execute(str_sql)

rens = []

for values in cur.fetchall():

    zhong_zhi = values[0][7:]
    ji_gou = values[1][11:]
    name = values[2][10:]
    ru_si_shi_jian = values[3][:10]
    che_bao_fei = bao_fei(name, zhong_zhi, '车险')
    cai_bao_fei = bao_fei(name, zhong_zhi, '财产险')
    ren_bao_fei = bao_fei(name, zhong_zhi, '人身险')

    rens.append((zhong_zhi,
                 ji_gou,
                 name,
                 ru_si_shi_jian,
                 che_bao_fei,
                 cai_bao_fei,
                 ren_bao_fei))

rens_sort = sorted(rens, key=lambda ren: ren[4], reverse = True)

for ren in rens_sort:
    print(ren)
