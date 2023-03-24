import urllib.parse
import os
import re
import time
import urllib.parse
import requests
from requests.exceptions import Timeout
from bs4 import BeautifulSoup

# path definition============================================================
path = os.getcwd()
dat_dir = os.path.join(path, 'dat')
csv_dir = os.path.join(path, r'dat/csv')
img_dir = os.path.join(path, r'dat/img')
output_dir = os.path.join(path, 'output')
conf_dir = os.path.join(path, 'config')
ex_temp_file = os.path.join(conf_dir, 'temp_file.xlsm')


# ===========================================================================
def scray_thumbnail(url):
    time.sleep(0.25)  # 高速すぎるので時間調整
    while True:
        try:
            html_source = requests.get(url, timeout=(3.0, 5.0))

            # BeautofulSoupが誤認識してしまうスクリプトを削除、これを除外すると読み込み不良になる
            # これを解除することでrequests.getによるデータ取得が可能になる
            html_source = html_source.text.replace('<script language="JavaScript" type="text/javascript">', '')

            soup = BeautifulSoup(html_source, 'lxml')
            flags = soup.find('img', src=re.compile('./指定されたページが見つかりません（エラー404）_ 楽天_files/w100.gif'))
            if flags:
                print('該当ページなし、スキップ')
                pass
            else:
                # tags_titles = soup.select('div.rnkRanking_itemName > a')
                # tags_titles = soup.select('#rnkRankingMain > div > div > '
                #                           'div.rnkRanking_detail > div > div > div > div.rnkRanking_itemName > a')
                # # tags_imgs = soup.select('div.rnkRanking_image > div > a > img')
                # tags_imgs = soup.select('#rnkRankingMain > div > div.rnkRanking_image > div > a > img')
                # return [(tit.text, img.attrs['src'], main4.filename_creation(img.attrs['src']),
                #          tit.attrs['href']) for tit, img in zip(tags_titles, tags_imgs) if tit]
                reviews_tags = soup.select('div.rnkRanking_starBox > div > a')
                return [(tag.getText for tag in reviews_tags)]
        except Timeout:
            # print('楽天サーバーの異常、処理を中断します')
            time.sleep(3)


# イメージファイル名生成
def filename_creation(src):
    return urllib.parse.urlparse(src)[2].replace('/', '')


if __name__ == '__main__':
    for page in range(1, 2):

        target_url = f'https://ranking.rakuten.co.jp/realtime/100371/p={page}'
        res = requests.get(target_url)
        if res.content:
            soup = BeautifulSoup(res.content, 'lxml')

        title_lists, title_urls, filenames, img_urls, revirews_lists, price_lists = [], [], [], [], [], []
        out_datas = []

        # get title,title_url,review
        tag_block = 'div.rnkRanking_upperbox'
        title_block_sources = soup.select(tag_block)
        # pprint(title_block_sources) # for test code
        for block_source in title_block_sources:
            # pprint(block_source.contents)

            # get litle,title_url
            title = block_source.select_one('div.rnkRanking_itemName > a')
            title_url = title.attrs['href']
            title_lists.append(title.text)
            title_urls.append(title_url)

            # get review
            review_tag = block_source.find('a', text=re.compile('レビュー'))
            if review_tag:
                # out = re.sub(r'\D', '', review_tag.text)
                revirews_lists.append(re.sub(r'\D', '', review_tag.text))
            else:
                revirews_lists.append('None')

        # get pcrice
        [price_lists.append(p.text.replace('円', '')) for p in soup.select('div.rnkRanking_price')]
        # get img
        [img_urls.append(img_url.attrs['src']) for img_url
         in soup.select('div.rnkRanking_image > div > a > img')]

        # todo:  entry ==  title, img_url, filename, title_url, price, review
        # todo: dataname == title_lists, title_urls, finames, img_urls, revirews_lists, price_list
        # todo: dataname == out_datas
        for i, ttl in enumerate(title_lists):
            out_datas.append([ttl, img_urls[i], '', title_urls[i], price_lists[i], revirews_lists[i]])
        print(out_datas[0])




        # todo:entry title, img_url, filename, title_url, price, review
        # print(price_lists)
        # print(title_lists)
        # print(title_urls)
        # print(img_urls[1])

if __name__ == '__main__':
    dt = ['a', 'b', 'c']
    print(dt)
    dt[2] = 'd'
    print(dt)
