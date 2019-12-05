import sqlite3


class Stats():

    def __init__(self, cur):
        self._zhong_zhi = None
        self._ji_gou = None
        self._name = None
        self._ru_si_shi_jian = None
        self._zhi_ji = None
        self._che = None
        self._cai = None
        self._ren = None
        self._cur = cur

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
    def zhi_ji(self):
        return self._zhi_ji

    @zhi_ji.setter
    def zhi_ji(self, str):
        self._zhi_ji = str

    def set_che(self, nian_fen=None, yue_fen=None):
        if nian_fen is not None:
            self._che = self.bao_fei('车险', f"{nian_fen}")
        elif yue_fen is not None:
            self._che = self.bao_fei('车险', f"{nian_fen}", f"{yue_fen}")
        else:
            self._che = self.bao_fei('车险')

    @property
    def che(self):
        return self._che

    def set_cai(self, nian_fen=None, yue_fen=None):
        if nian_fen is not None:
            self._cai = self.bao_fei('财产险', f"{nian_fen}")
        elif yue_fen is not None:
            self._cai = self.bao_fei('财产险', f"{nian_fen}", f"{yue_fen}")
        else:
            self._cai = self.bao_fei('财产险')

    @property
    def cai(self):
        return self._cai

    def set_ren(self, nian_fen=None, yue_fen=None):
        if nian_fen is not None:
            self._ren = self.bao_fei('人身险', f"{nian_fen}")
        elif yue_fen is not None:
            self._ren = self.bao_fei('人身险', f"{nian_fen}", f"{yue_fen}")
        else:
            self._ren = self.bao_fei('人身险')

    @property
    def ren(self):
        return self._ren

    @property
    def fei_che(self):
        return self.cai + self.ren

    @property
    def zheng_ti(self):
        return self.che + self.fei_che

    # def __repr__(self):
    #     return repr((self.zhong_zhi,
    #                  self.ji_gou,
    #                  self.name,
    #                  self.ru_si_shi_jian,
    #                  self.che_bao_fei,
    #                  self.cai_bao_fei,
    #                  self.ren_bao_fei))

    def bao_fei(self, xian_zhong, nian_fen=None, yue_fen=None):
        str_sql = f"SELECT SUM([签单保费/批改保费]) \
                    FROM 销售人员业务跟踪表 \
                    WHERE 业务员 like '%{self.name}%' \
                    AND 中心支公司 like '%{self.zhong_zhi}' \
                    AND [车险/财产险/人身险] = '{xian_zhong}'"

        if nian_fen is not None:
            str_sql += f" AND 年份 = '{nian_fen}'"

        if yue_fen is not None:
            str_sql += f" AND 月份 = '{yue_fen}'"

        self._cur.execute(str_sql)
        value = self._cur.fetchone()[0]
        if value is None:
            value = 0.0
        else:
            value /= 10000

        return value
