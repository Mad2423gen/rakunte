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

# todo: 追加分
# URLディレクトリ
id_dir = os.path.join(path, 'id')


# ===========================================================================
# 追加オーダー：idファイルを作成する（エクセルに網掛けを設定するため）
def save_id(genre, designated_column):
    csv_files = glob.glob(os.path.join(csv_dir, f'{genre}_*.csv'))
    output_file_name = os.path.join(id_dir, f'{genre}_id.txt')
    id_list = []
    for csv_file in csv_files:
        with open(csv_file, 'r', encoding='utf-8_sig', newline='') as rf:
            reader = csv.reader(rf)
            for row in reader:
                id_list.append(row[designated_column])
    if os.path.isfile(id_file):
        with open(output_file_name, 'a', encoding='utf-8_sig') as wf:
            for id_name in id_list:
                wf.write(str(id_name) + '\n')
    else:
        with open(output_file_name, 'w', encoding='utf-8_sig') as wf:
            for id_name in id_list:
                wf.write(str(id_name) + '\n')


# 返り値を集合で返す
def read_id_add_set():
    read_txts = set()
    text_list = glob.glob(os.path.join(id_dir, '*.txt'))
    for text in text_list:
        with open(text, 'r', encoding='utf-8_sig') as rf:
            for line in rf:
                read_txts.add(line.strip())
    pprint(read_txts)
    return read_txts

# ===========================================================================


if __name__ == '__main__':
    genre = 'realtime'
    id_file = os.path.join(id_dir, f'{genre}_id.txt')
    if os.path.isfile(id_file):
        os.remove(id_file)
    save_id(genre, 2)
    read_id_add_set()
