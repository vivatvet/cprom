#!/usr/bin/env python3

import pylightxl as xl
from typing import Tuple
import statistics
import math
import pprint


class ExcelFile:

    def __init__(self):
        self.title = []

    @staticmethod
    def load_table(file_name: str) -> Tuple[list, list]:
        db = xl.readxl(file_name)
        title = []
        raw_table = []
        first_row = True
        for row in db.ws(db.ws_names[0]).rows:
            if first_row:
                title = row
                first_row = False
                continue
            row = [str(v).strip() for v in row]
            raw_table.append(row)
        return title, raw_table

    # def group_by_kd(self, raw_table: list):
    #     tb = {}
    #     for row in raw_table:
    #         tb.setdefault(row[5], []).append(row)
    #     for k in tb.keys():
    #         print(k + " -- " + str(len(tb[k])))

    def group_by_npp(self, raw_table: list) -> dict:
        tb = {}
        for row in raw_table:
            work_row = self.string_to_float_item(row)
            tb.setdefault(row[0], []).append(work_row)
        return tb

    def string_to_float_item(self, row: list) -> list:
        m_row = []
        i = 0
        while i < len(row):
            if i == 8:
                m_row.append(float(row[i].replace(',', '.')))
            else:
                m_row.append(row[i])
            i += 1
        return m_row

    def sort_by_value(self, group_table: dict) -> dict:
        sorted_tabl = {}
        for group, row in group_table.items():
            sorted_row = sorted(row, key=lambda x: x[8], reverse=True)
            sorted_tabl[group] = sorted_row
        return sorted_tabl

    def get_90_cum_percent(self, group_table: dict) -> dict:
        tabl = {}
        for group, row in group_table.items():
            s = 0
            for r in row:
                s = s + r[8]
            previous_percent = 0
            for r in row:
                percent = (r[8] / s) * 100
                if previous_percent + percent <= 90:
                    tabl.setdefault(group, []).append(r)
                previous_percent = previous_percent + percent
        return tabl

    def stat_char(self, group_table: dict):
        average = {}
        sigma = {}
        chosen = {}
        not_chosen = {}
        step = 0
        average[step] = {}
        sigma[step] = {}
        for group, row in group_table.items():
            s = 0
            for r in row:
                s = s + r[8]
            average[step][group] = s / len(row)
            sig = 0
            for r in row:
                sig = sig + (r[8] - average[step][group]) * (r[8] - average[step][group])
            if len(row) >= 20 or len(row) == 1:
                sigma[step][group] = math.sqrt(sig / (len(row)))
            else:
                sigma[step][group] = math.sqrt(sig / (len(row) - 1))
            for r in row:
                if len(row) >= 20:
                    if r[8] >= (average[step][group] + (3 * sigma[step][group])):
                        chosen.setdefault(group, []).append(r)
                    else:
                        not_chosen.setdefault(group, []).append(r)
                elif len(row) <= 3:
                    chosen.setdefault(group, []).append(r)
                else:
                    if r[8] >= (average[step][group] + (2 * sigma[step][group])):
                        chosen.setdefault(group, []).append(r)
                    else:
                        not_chosen.setdefault(group, []).append(r)


