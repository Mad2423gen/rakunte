import csv
import os
import shutil
import time
import re
import pathlib
import datetime

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
# キーワードフォルダ
keyword_dir = os.path.join(conf_dir, 'keyword')
# タイムスタンプファイル
time_stamp_file = os.path.join(conf_dir, 'time_tamp.txt')
# キーワードファイル　共通.txt
common_keyword_file = os.path.join(keyword_dir, '共通.txt')

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
        ky_a = [line.rstrip("\n") for line in rf.readlines()]
    with open(os.path.join(keyword_dir, '共通.txt'), 'r', encoding='utf-8_sig') as rfb:
        ky_b = [line_b.rstrip("\n") for line_b in rfb.readlines()]
    return ky_a + ky_b


# ファイルの作成日付を取得、n日経過したら削除--------------------------------------------
def delete_old_files(n):
    if int(n) != 0:
        for filename in os.listdir(csv_dir):
            filepath = os.path.join(csv_dir, filename)
            if os.path.isfile(filepath) and time.time() \
                    - os.path.getctime(filepath) >= int(n) * 86400:
                print("保存期間経過、datファイル消去")
                [shutil.rmtree(rdr) for rdr in (csv_dir, img_dir)]
                # 時間計測ファイルも削除
                os.remove(time_stamp_file) if os.path.isfile(time_table_file) else None
                [os.makedirs(dr, exist_ok=True) for dr in (csv_dir, img_dir)]
                break


#  空のキーワードファイル作成(無条件)
def make_keyword_file():
    for lt in csv_read(genru_file):
        pathlib.Path(os.path.join(keyword_dir, f'{lt[0]}.txt')).touch()
    pathlib.Path(os.path.join(keyword_dir, '共通.txt')).touch()


# 不足しているキーワードディレクトリ、ファイルを作成
def make_keyword_file_missing():
    # ディレクトリ消失
    os.makedirs(keyword_dir) or make_keyword_file() if not os.path.isdir(keyword_dir) else None
    # ファイル消失
    file_name_lists_origin = set(f'{lt[0]}.txt' for lt in csv_read(genru_file))
    target_lists = set(os.listdir(keyword_dir))
    [print(f'キワードファイル補完：{replenish_file_name}')
     or pathlib.Path(os.path.join(keyword_dir, replenish_file_name)).touch()
     for replenish_file_name in file_name_lists_origin - target_lists
     if replenish_file_name]
    if not os.path.isfile(common_keyword_file):
        pathlib.Path(common_keyword_file).touch()


def url_duplicate_detection(save_data, intervaltime, genre):
    old_filename = f'{intervaltime}_{genre}.csv'
    if os.path.isfile(old_filename):
        if int(csv_read(time_table_file)[4][1]) == 1:
            old_datas = set(old_data[3] for old_data in csv_read(os.path.join(csv_dir, old_filename)))
            out = [dt for dt in save_data if dt[3] not in old_datas]
        else:
            out = save_data
    else:
        out = save_data
    return out


# import datetime
#
# specified_date = datetime.datetime(2022, 1, 1)  # 指定日を設定
#
# if datetime.datetime.now() > specified_date:
#     print("指定日を過ぎたため、プログラムを終了します。")
#     exit()
# else:
#     # 指定日以前の場合は、プログラムを続ける
#     # ここに続くプログラムを記述


# ---------------------------------------------------------------------------------
#
# if __name__ == '__main__':
#     savedata = ['https://item.rakuten.co.jp/book/17448477/?l2-id=Ranking_PC_realtime-101240-d_rnkRankingMain&s-id' \
#                 '=Ranking_PC_realtime-101240-d_rnkRankingMain_3',
#                 'https//item.rakuten-hogemoge'
#                 ]
#
#     output = url_duplicate_detection(savedata)
#     print(output)