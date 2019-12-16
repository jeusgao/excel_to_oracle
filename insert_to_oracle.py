#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2019-11-22 14:47:07
# @Author  : Joe Gao (jeusgao@163.com)
# @Link    : https://www.jianshu.com/u/3b77f85cc918
# @Version : $Id$

import xlrd
import glob
import pandas as pd
from tqdm import tqdm
import shutil
import uuid
import tools as pt

from config import *

import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

config_file = "config_oracle.yml"
conn = pt.open_connection(config_file)
cur = conn.cursor()

SRC_PATH = 'test'
OUTPUT_PATH = 'output'
fn_rangelog = 'out_of_range.log'
fn_missinglog = 'sheet_missing.log'

# if os.path.exists(fn_rangelog):
#     os.remove(fn_rangelog)
# if os.path.exists(fn_missinglog):
#     os.remove(fn_missinglog)
# if os.path.exists(OUTPUT_PATH):
#     shutil.rmtree(OUTPUT_PATH)

if not os.path.exists(OUTPUT_PATH):
    os.mkdir(OUTPUT_PATH)
out_sql = ''


def get_cell(tag, sheet, r, c):
    try:
        tmp = sheet.cell_value(r, c)
        tmp = tmp.replace('编制单位：', '').strip() if isinstance(tmp, str) else tmp
        if tmp:
            tmp = 0 if tmp in REPLACE_DIC else tmp
        else:
            tmp = ''
    except:
        with open(fn_rangelog, 'a') as f:
            f.write(f'Out of range: {tag} [{r}, {c}]\n')
        tmp = ''
    return tmp


def _get_field(p):
    tmp = []
    for n in list(p):
        if isinstance(n, str):
            tmp += [f"'{n}'"] if len(n) else ['NONE']
        else:
            tmp += [str(n)]

    return ', '.join(tmp)

fns = glob.glob(f'{SRC_PATH}/*.XLS')
for dic in CFG:
    headers = dic.get('head')
    dic_cols = dic.get('cols')
    sn = dic.get('sheet_name')
    fld_id = f'{sn.split(" ")[0]}ID'
    dropna_sub = dic.get('dropna_sub')
    print(f'{"_"*50}\nFORM {sn} Inserted ......')

    for fn in tqdm(fns):
        dic_out = {k[0]: [] for k in headers}
        dic_out[fld_id] = []
        wb = xlrd.open_workbook(filename=fn, formatting_info=True, logfile=open(os.devnull, 'w'))
        try:
            sheet = wb.sheet_by_name(sn)
        except:
            with open(fn_missinglog, 'a') as f:
                f.write(f'Sheet missing: {fn} ({sn})\n')
            continue

        max_len = 0
        for header in headers:
            h = header[0]
            col_sets = dic_cols.get(h)
            for dic_col in col_sets:
                tars = dic_col.get('tar_rows')
                col_type = dic_col.get('type')
                val = dic_col.get('val')
                if 'SINGLE_VALUE' in col_type:
                    if isinstance(tars, str):
                        dic_out[h] = [dic_col.get('val')] * max_len
                    else:
                        tmp_len = (max_len - len(dic_out.get(h, 0))) if tars < 0 else tars
                        dic_out[h] += [dic_col.get('val')] * tmp_len
                else:
                    if 'MULTI_CELL' in col_type:
                        if isinstance(val, list):
                            val[0][0] = list(range(val[0][0][0], sheet.nrows)) if isinstance(
                                val[0][0], tuple) else val[0][0]
                        dic_out[h] += [get_cell(f'{fn} ({sn}) -', sheet, row, pos[1]) for pos in val for row in pos[0]]
                    else:
                        dic_out[h] += [get_cell(f'{fn} ({sn}) -', sheet, val[0], val[1])] * max_len

            max_len = len(dic_out.get(h)) if len(dic_out.get(h)) > max_len else max_len

        dic_out[fld_id] = [f'{uuid.uuid1()}' for _ in range(max_len)]
        df = pd.DataFrame(dic_out)
        if dropna_sub:
            df = df[df.eval(dropna_sub[0]) != '']
        h_index = [fld_id] + [h[0] for h in sorted(headers, key=lambda x: x[1])]
        df = df[h_index]
        for row in df.iterrows():
            table = sn.split(' ')[0]
            _sql = f"insert into {table} ({', '.join(h_index)}) values ({_get_field(row[1])})".replace('NONE', "''")
            out_sql += f"{_sql};\n"
            cur.execute(_sql)
    conn.commit()

with open(f'{OUTPUT_PATH}/insert.sql', 'w') as f:
    f.write(f'{out_sql}')
pt.close_connection(conn)
