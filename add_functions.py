import csv
import os
import shutil
import time
import re
import requests
from bs4 import BeautifulSoup

import lxml

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


# 空のファイルを作る
def make_file(filename):
    with open(filename, 'w', encoding='utf-8_sig') as f:
        pass


# テキストファイルのキーワードを取得　リストで返す
def read_keywords(file):
    with open(file, 'r', encoding='utf-8_sig') as rf:
        ky_a = [line.rstrip("\n") for line in rf.readlines()]
    with open(common_keyword_file, 'r', encoding='utf-8_sig') as rfb:
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
                [shutil.rmtree(rdr, ignore_errors=True) for rdr in (csv_dir, img_dir)]
                # タイムスタンプファイルも削除
                os.remove(time_stamp_file) if os.path.isfile(time_table_file) else None
                # 削除したディレクトリ再生
                [os.makedirs(dr, exist_ok=True) for dr in (csv_dir, img_dir)]
                break


#  空のキーワードファイル作成(無条件)
def make_keyword_file():
    for lt in csv_read(genru_file):
        make_file(f'{lt[0]}.txt')
    make_file(common_keyword_file)


# 不足しているキーワードディレクトリ、ファイルを作成
def make_keyword_file_missing():
    # ディレクトリ消失
    os.makedirs(keyword_dir) or make_keyword_file() if not os.path.isdir(keyword_dir) else None
    # ファイル消失
    file_name_lists_origin = set(f'{lt[0]}.txt' for lt in csv_read(genru_file))
    target_lists = set(os.listdir(keyword_dir))
    [print(f'キワードファイル補完：{replenish_file_name}')
     or make_file(os.path.join(keyword_dir, replenish_file_name))
     for replenish_file_name in file_name_lists_origin - target_lists
     if replenish_file_name]
    if not os.path.isfile(common_keyword_file):
        make_file(common_keyword_file)


# 既に保存されているURLを新規データと比較して差分を返す
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


# レビューとプライスを付加したバージョン ※mainに移動したので一応使っていない　2023-03-13
def scray_thumbnail2(target_url):
    time.sleep(0.25)
    res = requests.get(target_url, timeout=(3.0, 5.0))
    html_source = res.text.replace('<script language="JavaScript" type="text/javascript">', '')

    # parse
    soup = BeautifulSoup(html_source, 'lxml')
    # Confirmation of page existence
    flags = soup.find('img', src=re.compile('./指定されたページが見つかりません（エラー404）_ 楽天_files/w100.gif'))

    if flags:
        pass
    else:
        # Declaration of tag element list
        title_lists, title_urls, filenames, img_urls, revirews_lists, price_lists = [], [], [], [], [], []
        # Declaration of elements for output
        out_datas = []

        # hint review_tagはあったりなかったりするので、先ず親タグからその部分のブロックを抽出、
        # 更にfindないしselectで抽出する、返り値がFalseの場合はレビューが存在しないと言うことになる。

        # get title,title_url,review
        tag_block = 'div.rnkRanking_upperbox'
        title_block_sources = soup.select(tag_block)
        # pprint(title_block_sources)    # For Developer Test
        for title_block_source in title_block_sources:
            # pprint(block_source.contents)   #For Developer Test

            # get litle,title_url
            title = title_block_source.select_one('div.rnkRanking_itemName > a')
            title_url = title.attrs['href']
            title_lists.append(str(title.text))
            title_urls.append(title_url)

            # get review  ※title_block内に「レビュー」が含まれなければNoneを返す
            review_tag = title_block_source.find('a', text=re.compile('レビュー'))
            if review_tag:
                revirews_lists.append(re.sub(r'\D', '', review_tag.text))
            else:
                revirews_lists.append('None')

        # get pcrice
        [price_lists.append(p.text.replace('円', '')) for p in soup.select('div.rnkRanking_price')]
        # get img
        [img_urls.append(img_url.attrs['src']) for img_url
         in soup.select('div.rnkRanking_image > div > a > img')]

        #  Reference Element List============================================================
        #   entry ==  title, img_url, filename, title_url, price, review
        #
        #  for out_datas
        #  dataname == title_lists, title_urls, finames, img_urls, revirews_lists, price_list
        #  dataname == out_datas
        # ===================================================================================

        for i, title in enumerate(title_lists):
            # entry is --> title, img_url, filename, title_url, price, review
            # filenameのエントリーは不要になったので空欄、他のロジックに影響するので削除しないこと
            out_datas.append((title, img_urls[i], '', title_urls[i], price_lists[i], revirews_lists[i]))
            # for developer testing -------------
            # print(price_lists)
            # print(title_lists)
            # print(title_urls)
            # print(img_urls[1])
            # -----------------------------------
        return out_datas
        # return print(out_datas[0]) # for developer testing


#
# ---------------------------------------------------------------------------------
#
if __name__ == '__main__':
    pass
