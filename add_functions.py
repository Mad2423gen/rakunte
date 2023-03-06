import csv
import os
import shutil
import time
import re
import pathlib

# path definition============================================================
path = os.getcwd()
dat_dir = os.path.join(path, 'dat')  # csc,imgフォルダの親フォルダ
csv_dir = os.path.join(dat_dir, 'csv')  # csv保存フォルダ
img_dir = os.path.join(dat_dir, 'img')  # img保存フォルダ
output_dir = os.path.join(path, 'output')  # 出力フォルダ
conf_dir = os.path.join(path, 'config')  # テンプレート、タイムテーブルファイル保存フォルダ
# エクセルテンプレートファイル
ex_temp_file = os.path.join(conf_dir, 'temp_file.xlsm')
# タイムテーブルファイル
time_table_file = os.path.join(conf_dir, 'timetable.csv')
# 除外キーワードリスト
exclusion_keyword_file = os.path.join(conf_dir, 'exclusion_keyword.txt')
# ジャンルcsv
genru_file = os.path.join(conf_dir, 'rakuten_genre.csv')
# ジャンルフォルダ
keyword_dir = os.path.join(conf_dir, 'keyword')


# ===========================================================================
# csv read タイトルのみ取得　リストで返す
def csv_read_title(csv_file):
    return [row[0] for row in csv.reader(open(csv_file, 'r', encoding='utf-8_sig', newline=''))]


# csvを読み込んでリストで返すだけ
def csv_read(csv_file):
    return [row for row in csv.reader(open(csv_file, 'r', encoding='utf-8_sig', newline=''))]


# テキストファイルのキーワードを取得　リストで返す
def read_keywords(file):
    with open(file, 'r', encoding='utf-8_sig') as rf:
        return [line.rstrip("\n") for line in rf.readlines()]


# ファイルの作成日付を取得、n日経過したら削除--------------------------------------------
def delete_old_files(n):
    if int(n) != 0:
        for filename in os.listdir(csv_dir):
            filepath = os.path.join(csv_dir, filename)
            if os.path.isfile(filepath) and time.time() \
                    - os.path.getctime(filepath) >= int(n) * 86400:
                print("保存期間経過、datファイル消去")
                shutil.rmtree(csv_dir)
                shutil.rmtree(img_dir)
                os.makedirs(csv_dir, exist_ok=True)
                os.makedirs(img_dir, exist_ok=True)
                break


#  空のキーワードファイル作成(無条件)
def make_keyword_file():
    [pathlib.Path(os.path.join(keyword_dir, f'{lt[0]}.txt')).touch() for lt in csv_read(genru_file)]


# 不足しているキーワードディレクトリ、ファイルを作成
def make_keyword_file_missing():
    # ディレクトリ消失
    os.makedirs(keyword_dir) or make_keyword_file() or \
    print('消失したキーワードディレクトリ＆ファイルを作成') if not os.path.isdir(keyword_dir) else None
    # ファイル消失
    file_name_lists_origin = set(f'{lt[0]}.txt' for lt in csv_read(genru_file))
    target_lists = set(os.listdir(keyword_dir))
    [print(f'キワードファイル補完：{replenish_file_name}')
     or pathlib.Path(os.path.join(keyword_dir, replenish_file_name)).touch()
     for replenish_file_name in file_name_lists_origin - target_lists
     if replenish_file_name]


# ---------------------------------------------------------------------------------

if __name__ == '__main__':
    make_keyword_file_missing()
