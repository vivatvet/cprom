#!/usr/bin/env python3

import pylightxl as xl
from typing import Tuple
import statistics
import math
import pprint
import queue


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

    def get_stat_char(self, row: list) -> Tuple[float, float, int, float]:
        s = 0
        for r in row:
            s = s + r[8]
        average = s / len(row)
        sig = 0
        for r in row:
            sig = sig + (r[8] - average) * (r[8] - average)
        if len(row) >= 20 or len(row) == 1:
            sigma = math.sqrt(sig / (len(row)))
        else:
            sigma = math.sqrt(sig / (len(row) - 1))
        count = len(row)
        covar = (sigma / average) * 100
        return average, sigma, count, covar

    def divide_row(self, row: list, average) -> Tuple[list, list]:
        ll1 = []
        ll2 = []
        for r in row:
            if r[8] > average:
                ll1.append(r)
            else:
                ll2.append(r)
        return ll1, ll2

    def stat_char(self, group_table: dict):
        average = {}
        sigma = {}
        chosen = {}
        strata = {}
        not_chosen = {}
        step = 0
        average[step] = {}
        sigma[step] = {}
        q = queue.Queue()
        for group, row in group_table.items():
            average[step][group], sigma[step][group], count, _ = self.get_stat_char(row)
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
        for group, row in not_chosen.items():
            step = 1
            q.put(row)
            while not q.empty():
                average[step] = {}
                sigma[step] = {}
                rq = q.get()
                average[step][group], sigma[step][group], count, covar = self.get_stat_char(rq)
                if covar > 33:
                    row_1, row_2 = self.divide_row(rq, average[step][group])
                    q.put(row_1)
                    q.put(row_2)
                else:
                    strata.setdefault(group, []).append(rq)
                step += 1
