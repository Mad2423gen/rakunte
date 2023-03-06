import csv
import glob
import shutil
import os
import urllib.parse
from datetime import datetime

# path definition============================================================
path = os.getcwd()
dat_dir = os.path.join(path, 'dat')
csv_dir = os.path.join(path, r'dat/csv')
img_dir = os.path.join(path, r'dat/img')
output_dir = os.path.join(path, 'output')
conf_dir = os.path.join(path, 'config')
ex_temp_file = os.path.join(conf_dir, 'temp_file.xlsm')
# ===========================================================================

# 現在時刻
# def add_datetime():
#     return datetime.now().strftime('%Y-%m-%d-%H_%M_%S')
#
# intervaltime = 'realtime'
#
# output_ex_files_dir = os.path.join(output_dir, intervaltime, add_datetime())
# shutil.move(img_dir, os.path.join(output_ex_files_dir, 'img'))

# file_names = [os.path.basename(file_path) for file_path in glob.glob(os.path.join(csv_dir, '*.csv'))]
# print(file_names)
# names = [os.path.splitext(name)[0] for name in file_names]
# print(names)
"""
def filename_creation(src):
    return urllib.parse.urlparse(src)[2].replace('/', '')

# csv read
def read_csv(csvfile='レディースファッション.csv'):
    with open(os.path.join(csv_dir, csvfile),
              'r', encoding='utf-8_sig', newline='') as tf:
        return [filename_creation(row[1]) for row in csv.reader(tf)]

# print(old_img_filelists)

new_img_filelists = ['sample1.jpg',
                     'photo2.jpg',
                     'nature_scene.jpg',
                     'sunset_photo.jpg',
                     'family_pic.jpg',
                     'denimlifecabinetscca65ca6612_2023th.jpg'
                     ]

# Find keywords not included
for new_img_filelist in new_img_filelists:
    print(f"{new_img_filelist} is in the list.\n"
          if {new_img_filelist}.issubset(set(read_csv()))
          else f"{new_img_filelist} is not in the list.\n")
"""

# -----------------------------------------------------------------------------


"""
リストから読み込んだキーワードをpattern = re.compile()の引数に代入するコードを書く

リストから読み込んだキーワードを正規表現パターンに含めるには、
リストの要素を正規表現のOR条件に変換してパターンを作成する必要があります。
以下はPythonでリストから読み込んだキーワードを正規表現パターンに含めるコードの例です。
"""
import re

# 検索対象のキーワードリスト
keywords = ['PS5', 'Nintendo', 'FPS']
# リストの要素を正規表現のOR条件に変換して正規表現パターンを作成
pattern = re.compile('|'.join(keywords))
# re.compile('PS5|Nintendo|FPS')

# テスト用の文字列
text = 'I like to play FPS games on my Nintendo Switch, ' \
       'but I am thinking of getting a PS5.'

# テキストからキーワードを検索
if pattern.search(text):
    print('Found a keyword!')
else:
    print('Keyword not found.')

"""
この例では、リストの要素を正規表現のOR条件に変換して、'PS5'、
'Nintendo'、または'FPS'のいずれかを含むキーワードを検索する
正規表現パターンを作成しています。そして、テキストからキーワードを検索し、
'Found a keyword!'という出力がされます。
"""
# -----------------------------------------------------------------------------

# a = ['apple','orange','grape']　に　b = ['itary','japan','usa']を結合するコードは？
# リストの結合には、+演算子を使用することができます。
# 以下はPythonでリストaとbを結合するコードの例です。
a = ['apple', 'orange', 'grape']
# a = []
b = ['Italy', 'Japan', 'USA']
c = a + b
print(f'|'.join(c))
# ['apple', 'orange', 'grape', 'Italy', 'Japan', 'USA']
# -----------------------------------------------------------------------------

