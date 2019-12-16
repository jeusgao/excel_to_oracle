#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2019-11-22 14:47:07
# @Author  : Joe Gao (jeusgao@163.com)
# @Link    : https://www.jianshu.com/u/3b77f85cc918
# @Version : $Id$

import os

YEAR = 2018
MONTH = 12
REF_DIC = {'income': '收入', 'cost': '支出',
           'xzdw': '行政单位', 'sydw': 'XX单位', 'qyhgl': 'XXXXX事业单位', 'mjfyl': 'XXXXX组织',
           'sgjf': 'XXXX支出', 'jgyx': 'XXXX经费', 'gyzczy': 'XX资产占用情况',
           'zwslgqk': 'XXXXX情况（年末数）', 'zwjggdzc': 'XXXX固定资产（年末数）', 'zwjgzczl': 'XXXX资产总量情况（年末数）', 'zwjgjfzc': 'XXXX经费支出明细（本年数）'}
REPLACE_DIC = ['—', '-']

CFG = [
    {'sheet_name': 'Z01 XXXXXXXX',
        'head': [('XM', 3), ('AMTTYPE', 4), ('AMT0', 5), ('AMT1', 6), ('AMT2', 7), ('AGENCY_CODE', 0), ('SET_YEAR', 1), ('SET_MONTH', 2)],
        'cols':
        {
            'XM': [{'tar_rows': 'ALL',
                    'type': 'MULTI_CELL',
                    'val': [[list(range(6, 13)), 0], [list(range(6, 29)), 5], [list(range(6, 15)), 10], [list(range(16, 28)), 10]]}],
            'AMTTYPE': [{'tar_rows': len(list(range(6, 13))),
                         'type': 'SINGLE_VALUE',
                         'val': REF_DIC.get('income')},
                        {'tar_rows': -1,
                         'type': 'SINGLE_VALUE',
                         'val': REF_DIC.get('cost')}],
            'AMT0': [{'tar_rows': 'ALL',
                      'type': 'MULTI_CELL',
                      'val': [[list(range(6, 13)), 2], [list(range(6, 29)), 7], [list(range(6, 15)), 12], [list(range(16, 28)), 12]]}],
            'AMT1': [{'tar_rows': 'ALL',
                      'type': 'MULTI_CELL',
                      'val': [[list(range(6, 13)), 3], [list(range(6, 29)), 8], [list(range(6, 15)), 13], [list(range(16, 28)), 13]]}],
            'AMT2': [{'tar_rows': 'ALL',
                      'type': 'MULTI_CELL',
                      'val': [[list(range(6, 13)), 4], [list(range(6, 29)), 9], [list(range(6, 15)), 14], [list(range(16, 28)), 14]]}],
            'AGENCY_CODE': [{'tar_rows': 'ALL', 'val': [2, 0], 'type': 'SINGLE_CELL'}],
            'SET_YEAR': [{'tar_rows': 'ALL', 'val': YEAR, 'type': 'SINGLE_VALUE'}],
            'SET_MONTH': [{'tar_rows': 'ALL', 'val': MONTH, 'type': 'SINGLE_VALUE'}],
        }
     },
]
