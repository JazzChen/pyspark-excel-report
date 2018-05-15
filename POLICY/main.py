#!/usr/bin/env python
# -*- coding:utf-8 -*-

# unicode setting for chinese characters
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
# python libs
import numpy as np
import pandas as pd
# pyspark libs
from pyspark import SparkConf, SparkContext
from pyspark.sql import HiveContext
# plot libs
from utils.plot import *

# 读取Hive表记录
jobname="risk_strategy_monited_report"
conf = SparkConf()
conf.setAppName(jobname)
sc = SparkContext(conf=conf)
hiveCtx = HiveContext(sc)
# log
log4jLogger = sc._jvm.org.apache.log4j
LOGGER = log4jLogger.LogManager.getLogger(jobname)

def get_pandas_from_hive(sql_text):
    df = hiveCtx.sql(sql_text)
#    df.cache()
    return df.toPandas()

# export results to excel
import datetime
import os
import json
import codecs

today = datetime.datetime.today()
dest_filename = u'$HOME/POLICY/report/数据监控日报{}月{}日.xlsx'.format(today.month, today.day)
resource_path = u'$HOME/POLICY/resource'

sheets = []
for d in os.listdir(resource_path):
    if os.path.isdir(os.path.join(resource_path, d)):
#       sheets.append(d.decode('utf-8'))
        sheets.append(d)

wb = Workbook()
ws = wb.active
ws.title = sheets[0]

for s in sheets[1:]:
    wb.create_sheet(title=s)

for s in sheets:
    LOGGER.info(s)
    ws = wb[s]
    file_paths = []
    sheet_path = os.path.join(resource_path, s)
    for file_name in os.listdir(sheet_path):
         file_path = os.path.join(sheet_path, file_name)
         if os.path.isfile(file_path) and os.path.splitext(file_name)[1] == '.json':
              file_paths.append(file_path)

    current_row = 2
    current_column = 2
    for file_path in sorted(file_paths):
        js = {}
        with codecs.open(file_path, 'r', 'utf-8') as json_file:
            js = json.load(json_file)

        LOGGER.info(file_path)
        sql_text = js['sql']
        LOGGER.info(sql_text)
        index = js['index']
        index_name = js['index_name']
        title = js['title']
        df = get_pandas_from_hive(sql_text)
        du = df.set_index(index)        
        # title, header, body,blank lines
        rows = len(du.index.levels[-1]) + 5
        start_cell = tuple_to_coordinate(current_row, current_column)
        plot_table(ws, du, start_cell, title, index_name)
        current_row = current_row + rows

    # auto adjust column wide
    for col in ws.columns:
        max_length = 4
        column = col[0].column # Get the column name
        for cell in col:
            if cell.coordinate in ws.merged_cells:
                continue
            if type(cell.value) != type(None):
                length = len(str('%.2f' % cell.value)) if isinstance(cell.value, np.float64) else len(str(cell.value))
                if length > max_length:
                    max_length = length
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width

wb.save(dest_filename)

