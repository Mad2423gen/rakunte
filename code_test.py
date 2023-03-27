import csv
import glob
import os
import sys
import random
import re
import shutil
import time
from pprint import pprint

import lxml
import subprocess
import urllib.parse
from datetime import datetime
from functools import wraps
import requests
from requests.exceptions import Timeout
import schedule
import win32com.client
from bs4 import BeautifulSoup
from joblib import Parallel, delayed
import add_functions
import logging

# path definition============================================================
path = os.getcwd()
dat_dir = os.path.join(path, 'dat')  # csc,imgフォルダの親フォルダ
csv_dir = os.path.join(dat_dir, 'csv')  # csv保存フォルダ
img_dir = os.path.join(dat_dir, 'img')  # img保存フォルダ
output_dir = os.path.join(path, 'output')  # 出力フォルダ
conf_dir = os.path.join(path, 'config')  # テンプレート、タイムテーブルファイル保存フォルダ
keyword_dir = os.path.join(conf_dir, 'keyword')  # 除外キーワードディレクトリ
# エクセルテンプレートファイル
ex_temp_file = os.path.join(conf_dir, 'temp_file.xlsm')
# タイムテーブルファイル
time_table_file = os.path.join(conf_dir, 'timetable.csv')
# タイムスタンプファイル
time_stamp_file = os.path.join(conf_dir, 'time_tamp.txt')
# キーワードファイル　共通.txt
common_keyword_file = os.path.join(keyword_dir, '共通.txt')
# URLディレクトリ
id_dir = os.path.join(path, 'id')


# ===========================================================================
# 追加オーダー：idファイルを作成する（エクセルに網掛けを設定するため）
# 第一引数はインポートするジャンル名、第二引数はid要素の列番
def save_id(time_interval, designated_column):
    csv_files = glob.glob(os.path.join(csv_dir, f'{time_interval}_*.csv'))
    id_list = []
    for csv_file in csv_files:
        with open(csv_file, 'r', encoding='utf-8_sig', newline='') as rf:
            reader = csv.reader(rf)
            for row in reader:
                id_list.append(row[designated_column])
    save_filename = os.path.join(id_dir, f'{time_interval}_id.txt')
    with open(save_filename, 'a', encoding='utf-8_sig') as wf:
        for row in id_list:
            print(row)
            wf.write(row + '\n')


# idファイルをインポートして、set型に変換する、引数はジャンル名
def read_id_add_set(time_interval):
    read_txt = set()
    read_filename = os.path.join(id_dir, f'{time_interval}_id.txt')
    with open(read_filename, 'r', encoding='utf-8_sig') as rf:
        for row in rf:
            read_txt.add(row.strip())
    return read_txt

# ===========================================================================


if __name__ == '__main__':
    # id_files = glob.glob(os.path.join(id_dir, '*.txt'))
    # if id_files:
    #     for id_file in id_files:
    #         os.remove(id_file)

    time_interval = 'realtime'
    designated_column = 2
    # save_id(time_interval, designated_column)
    s_list = read_id_add_set(time_interval)
