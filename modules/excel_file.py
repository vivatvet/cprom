#!/usr/bin/env python3

import pylightxl as xl
from typing import Tuple
import math
import queue
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
        for rows in raw_table:
            work_rows = self.string_to_float_item(rows)
            tb.setdefault(rows[0], []).append(work_rows)
        return tb

    @staticmethod
    def string_to_float_item(rows: list) -> list:
        m_rows = []
        i = 0
        while i < len(rows):
            if i == 8:
                m_rows.append(float(rows[i].replace(',', '.')))
            else:
                m_rows.append(rows[i])
            i += 1
        return m_rows

    @staticmethod
    def sort_by_value(group_table: dict) -> dict:
        sorted_table = {}
        for group, rows in group_table.items():
            sorted_rows = sorted(rows, key=lambda x: x[8], reverse=True)
            sorted_table[group] = sorted_rows
        return sorted_table

    @staticmethod
    def get_90_cum_percent(group_table: dict) -> Tuple[dict, dict]:
        table = {}
        not_selected_table = {}
        for group, row in group_table.items():
            s = 0
            for r in row:
                s = s + r[8]
            previous_percent = 0
            for r in row:
                percent = (r[8] / s) * 100
                if previous_percent + percent <= 90:
                    table.setdefault(group, []).append(r)
                else:
                    not_selected_table.setdefault(group, []).append(r)
                previous_percent = previous_percent + percent
        return table, not_selected_table

    @staticmethod
    def get_stat_char(rows: list) -> Tuple[float, float, int, float]:
        s = 0
        for r in rows:
            s = s + r[8]
        average = s / len(rows)
        sig = 0
        for r in rows:
            sig = sig + (r[8] - average) * (r[8] - average)
        if len(rows) >= 20 or len(rows) == 1:
            sigma = math.sqrt(sig / (len(rows)))
        else:
            sigma = math.sqrt(sig / (len(rows) - 1))
        count = len(rows)
        covar = (sigma / average) * 100
        return average, sigma, count, covar

    @staticmethod
    def divide_rows(rows: list, average) -> Tuple[list, list]:
        rows_1 = []
        rows_2 = []
        for r in rows:
            if r[8] > average:
                rows_1.append(r)
            else:
                rows_2.append(r)
        return rows_1, rows_2

    def stat_char(self, group_table: dict) -> Tuple[dict, dict, dict, dict, dict, dict]:
        average = {}
        sigma = {}
        count = {}
        covar = {}
        chosen = {}
        strata = {}
        not_chosen = {}
        step = 0
        q = queue.Queue()
        for group, rows in group_table.items():
            average[group] = {}
            sigma[group] = {}
            count[group] = {}
            average[group][step], sigma[group][step], count[group][step], _ = self.get_stat_char(rows)
            for r in rows:
                if len(rows) >= 20:
                    if r[8] >= (average[group][step] + (3 * sigma[group][step])):
                        chosen.setdefault(group, []).append(r)
                    else:
                        not_chosen.setdefault(group, []).append(r)
                elif len(rows) <= 3:
                    chosen.setdefault(group, []).append(r)
                else:
                    if r[8] >= (average[group][step] + (2 * sigma[group][step])):
                        chosen.setdefault(group, []).append(r)
                    else:
                        not_chosen.setdefault(group, []).append(r)
        for group, rows in not_chosen.items():
            average[group] = {}
            sigma[group] = {}
            count[group] = {}
            covar[group] = {}
            step = 1
            q.put(rows)
            while not q.empty():
                rows_q = q.get()
                average[group][step], sigma[group][step], count[group][step], covar[group][step] = self.get_stat_char(rows_q)
                if covar[group][step] > 33:
                    row_1, row_2 = self.divide_rows(rows_q, average[group][step])
                    q.put(row_1)
                    q.put(row_2)
                else:
                    strata.setdefault(group, []).append(rows_q)
                step += 1
        return chosen, strata, average, sigma, count, covar

    def random_choose(self, strata: dict):
        for group, strata_item in strata.items():
            dn = 0
            count_all = 0
            sum_group = 0
            for i, rows in enumerate(strata_item):
                _, sigma, count, _ = self.get_stat_char(rows)
                dn = dn + (sigma * sigma * count)
                count_all = count_all + count
                for r in rows:
                    sum_group = sum_group + r[8]
            variance_group = dn / count_all
            # opt_number =





    def write_to_file(self, file_name: str, tb_by_npp: dict, tb_title: list, chosen: dict):
        db = xl.Database()
        sheet = 'Processed'
        db.add_ws(sheetname=sheet, data={})
        for col_id, data in enumerate(tb_title, start=1):
            db.ws(sheet).update_index(row=1, col=col_id, val=data)
        i = 2
        for group in tb_by_npp.keys():
            # chosen
            if chosen.get(group):
                for rows in chosen[group]:
                    for col_id, data in enumerate(rows + ['BIG'], start=1):
                        db.ws(sheet).update_index(row=i, col=col_id, val=data)
                    i += 1
                i += 1
        xl.writexl(db, file_name)

    def is_file_exist(self, file_name):
        pass
