import sqlite3


class Stats():

    def __init__(self, name, table_name):
        self.name = name
        self.t_name = table_name

    @property
    def zheng_ti_bao_fei(self):
        str_sql = f"SELECT "
