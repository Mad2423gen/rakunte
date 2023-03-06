import csv
import glob
import os
import random
import shutil
import time
import lxml
import subprocess
import urllib.parse
from datetime import datetime
from functools import wraps
import requests
import schedule
import win32com.client
from bs4 import BeautifulSoup
from joblib import Parallel, delayed
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import add_functions

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
# ===========================================================================
# Decorator for retries
def retry(max_attempts, wait_time):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            attempts = 0
            while True:
                try:
                    result = func(*args, **kwargs)
                    return result
                except:
                    attempts += 1
                    if attempts >= max_attempts:
                        raise
                    time.sleep(wait_time)

        return wrapper

    return decorator

# ===========================================================================
# 現在日時取得 get Current date and time
def add_datetime():
    return datetime.now().strftime('%Y-%m-%d-%H_%M_%S')

# ===========================================================================
# イメージファイル名生成
def filename_creation(src):
    return urllib.parse.urlparse(src)[2].replace('/', '')

# ===========================================================================
# image download
# When there is only one URL to link to
def img_save(src, save_dir=img_dir):
    # if not os.path.isdir(save_dir):
    #     os.mkdir(save_dir)
    os.makedirs(save_dir, exist_ok=True)
    while True:
        try:
            with open(os.path.join(save_dir, filename_creation(src)), "wb") as f:
                f.write(requests.get(src).content)
            break
        except requests.exceptions.RequestException as e:
            print(e)
            print('***ダウンロードエラー リトライします***')
            print('*******Download error, retry*******')
            time.sleep(1)

# ===========================================================================
# If the linked URL is listed (parallel processing)
# ※複数のURLから画像をダウンロードするための並列処理
def img_saves(urls):
    # (withブロックを抜けるまで待機する)　n_jabs=n がスレッド数 verbose=1はメッセージ深度
    with Parallel(n_jobs=-1, verbose=1) as parallel:
        parallel(delayed(img_save)(url) for url in urls)

# ===========================================================================
# エクセルにエクスポート
# 引数： output_ex_files_dir == otput/time_interval_dir/datetime
def export_ex(output_ex_files_dir, intervaltime):
    # csvファイルリスト作成
    csv_filenames = [os.path.basename(file_path)
                     for file_path in glob.glob(os.path.join(csv_dir, f'{intervaltime}_*.csv'))]

    # error防止: 残っているエクセルタスクを強制終了
    subprocess.run('taskkill /F /T /IM excel.exe', stdout=None, shell=True)

    # rakunte/output/intervaltime/日時フォルダに移動しての処理-------------------
    os.chdir(output_ex_files_dir)
    for csv_file in csv_filenames:
        # outputfilename = f"{os.path.splitext(csv_file)[0]}.xlsm"
        # csvから画像ファイル名抽出、ファイル名抽出
        print(f'{csv_file}をエクスポート')

        # 作業用ファイルのダミーファイル
        dummy_filename = fr'{random.randint(0, 100000)}.xlsm'
        # フルパス
        dummy_file_path = os.path.join(output_ex_files_dir, dummy_filename)

        # エクセルのテンプレートファイルをダミーファイル名でコピー（重複防止）
        shutil.copyfile(ex_temp_file, dummy_file_path)

        # from csv to write exel
        with open(os.path.join(csv_dir, csv_file),
                  'r', encoding='utf-8_sig', newline='') as csvf:
            reader = csv.reader(csvf)
            try:
                # win32com
                excel = win32com.client.Dispatch('Excel.Application')
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(dummy_file_path)
                sheet = wb.Worksheets('Sheet1')
                sheet.Activate()
                print('excel writing')
                for i, lne in enumerate(reader):
                    sheet.Cells(i + 2, 1).Value = lne[2]
                    sheet.Cells(i + 2, 4).Value = lne[0]
                    sheet.Cells(i + 2, 5).Value = lne[3]
                    cell = sheet.Cells(i + 2, 5)
                    cell.Hyperlinks.Add(cell, lne[3])
                print('writing termination')
                # wb.Save()
                # wb.Close()
            except:
                print('!!!With error exel-write-handling!!!')

            # time.sleep(1)

            try:
                print('vba start')
                # xl2 = win32com.client.Dispatch('Excel.Application')
                # xl2.Workbooks.Open(dummy_file_path)
                excel.Application.Run('Module1.getimg')
                excel.Workbooks(1).Close(SaveChanges=1)
                print('vba termination')
            except:
                print('!!!With error vba-handling!!!')

        # time.sleep(0.5)

        # run vba
        # try:
        #     print('vba start')
        #     xl2 = win32com.client.Dispatch('Excel.Application')
        #     xl2.Workbooks.Open(dummy_file_path)
        #     xl2.Application.Run('Module1.getimg')
        #     xl2.Workbooks(1).Close(SaveChanges=1)
        #     print('vba termination')
        # except:
        #     print('!!!With error vba-handling!!!')

        time.sleep(1)

        # rename ダミーファイル名をジャンル名に
        try:
            os.rename(dummy_filename, f'{os.path.splitext(csv_file)[0]}.xlsm')
            print('rename termination\n')
        except:
            print('!!!With error rename-handling!!!')

    # カレントディレクトリに復帰
    os.chdir(path)

# ===========================================================================
# Page source acquisition block
# スクレイピング本体、　１ページのソースからタグを検出、データ取得
@retry(max_attempts=5, wait_time=10)  # 最大５回、10秒後リトライ
def scray_thumbnail(url, driver):
    driver.get(url)
    # 要素が読み込まれるまで待機するためのダミーメソッド
    dummy_tags = driver.find_elements(By.CLASS_NAME, 'rnkRanking_itemName')
    # ページソース取得
    html_source = driver.page_source
    # 保存するデータ抽出(use BeautifulSoup4)
    soup = BeautifulSoup(html_source, 'lxml')  # importでlxmlを削除するとエラーになる
    tags_titles = soup.select('div.rnkRanking_itemName > a')
    tags_imgs = soup.select('div.rnkRanking_image > div > a > img')

    return [(tit.text, img.attrs['src'], filename_creation(img.attrs['src']),
             tit.attrs['href']) for tit, img in zip(tags_titles, tags_imgs) if tit]

# ===========================================================================
# スクレイピングとCSV保存、画像保存
def csv_save(genre, genre_id, intervaltime, driver):
    # スクレイピング　（ジャンル内の全ページデータ取得）
    global old_csv_datas, keywords, exclusion_keywords, save_data
    print('\nrakuten scray')
    new_data = []
    for i in range(1, 6):  # 1~５ページ
        print(f'page:{i}')
        # target_url
        url = f'https://ranking.rakuten.co.jp/{intervaltime}/{genre_id}/p={i}'

        # スクレイピングしたデータを new_data に格納
        [new_data.append([ttl[0], ttl[1], ttl[2], ttl[3]]) for ttl in scray_thumbnail(url, driver)]

    # =====================csv保存データ作成==================================

    # キーワードディレクトリ＆ファイル確認(消失していた場合は補完）
    add_functions.make_keyword_file_missing()
    # ジャンル別キーワードファイル
    exclusion_keyword_file = os.path.join(keyword_dir, f'{genre}.txt')

    # 古いジャンルcsvファイル無
    if not os.path.isfile(os.path.join(csv_dir, f'{intervaltime}_{genre}.csv')):
        print('ジャンルCSV無、', end='')
        # 除外キーワードファイル有
        if os.path.isfile(exclusion_keyword_file):
            print('除外キーワードファイル有、', end='')
            keywords = add_functions.read_keywords(exclusion_keyword_file)
            # キーワード登録無
            if not keywords:
                # new_dataをそのまま保存
                print('キーワード登録無、新規データをジャンルcsvへ保存')
                save_data = new_data
            # キーワード登録有
            else:
                print('キーワード登録有、該当を削除、ジャンルCSVへ保存')
                # new_dataからキーワード対象を除外する
                save_data = [x for x in new_data if not any(y in x[0] for y in keywords)]

        # 除外キーワードファイル無
        else:
            # new_dataをそのまま保存
            print('除外キーワードファイル無、新規データをそのまま保存')
            save_data = new_data

    # 古いジャンルcsvファイル有
    else:
        print('ジャンルCSV有、', end='')

        # 除外キーワードファイル無
        if not os.path.isfile(exclusion_keyword_file):
            # 古いジャンルcsvと新規データを比較、差分をnew_dataへ保存
            print('キーワードファイル無、新旧ジャンルcsvの差分をジャンルcsvへ保存')
            old_data = add_functions\
                .csv_read_title(os.path.join(csv_dir, f'{intervaltime}_{genre}.csv'))
            save_data = [nd for nd in new_data if not nd[0] in old_data]

        # 除外キーワードファイル有
        else:
            # キーワード登録無し
            keywords = add_functions.read_keywords(exclusion_keyword_file)
            print('除外キーワードファイル有、', end='')
            if not keywords:
                # 古いジャンルcsvと新規データを比較、差分をnew_dataへ保存
                print('キーワード登録無、新旧データを比較、差分を保存')
                old_data = add_functions\
                    .csv_read_title(os.path.join(csv_dir, f'{intervaltime}_{genre}.csv'))
                save_data = [nd for nd in new_data if not nd[0] in old_data]

            # キーワード登録有り
            else:
                # ジャンルcsvとキーワードを結合、new_dataと比較する
                print('キーワード登録有、登録キーワードを結合、新データと比較、差分を保存')
                old_data = add_functions\
                    .csv_read_title(os.path.join(csv_dir, f'{intervaltime}_{genre}.csv'))
                joint_data = old_data + keywords
                save_data = [x for x in new_data if not any(y in x[0] for y in joint_data)]

    # img&csv保存===========================================================
    # save_dataから画像のリンクを取得し、ダウンロード（並列処理）
    print('img save now')
    img_saves([row[1] for row in save_data])
    print('img saved', '\n')

    # csvへ取得データ保存
    print('csv save now')
    with open(os.path.join(csv_dir, f'{intervaltime}_{genre}.csv'),
              'a', encoding='utf-8_sig', newline='') as sdf:
        csv.writer(sdf).writerows(save_data)
    print('csv saved\n')

# ===========================================================================
# 全処理実行関数　mode:リアルタイム、デイリー、ウィークリー選択　　mode2:テスト、本番実行選択
def main_func(mode=1, mode2=1):
    # 開始時刻取得
    start_time = add_datetime()
    print('==========スクレイピング処理開始==========')

    # counting period
    global intervaltime, genre_file
    intervaltime = 'realtime' if mode == 1 else 'daily' \
        if mode == 2 else 'weekly' if mode == 3 else 'unknown'

    print(f'スクレイピング範囲：{intervaltime}')
    # path definition----------------------------------------------
    # 取得ジャンル一覧を読み込み
    if mode2 == 1:  # テスト用
        print('\n=====debug mode=====\n')
        genre_file = os.path.join(path, r'config/test_rakuten_genre.csv')
    elif mode2 == 2:  # 本実行
        print('\n=====main function=====\n')
        genre_file = os.path.join(path, r'config/rakuten_genre.csv')
    # path definition----------------------------------------------

    # csvファイル更新日確認(datフォルダ消去、作成)
    print('Check update interval of csv files')

    # 新タイミングを読み込み、設定に従って削除する　
    # 設定日数取得
    specified_date = add_functions.csv_read(time_table_file)[3][1]

    # フォルダ作成、存在する場合はスキップ
    os.makedirs(csv_dir, exist_ok=True)
    os.makedirs(img_dir, exist_ok=True)
    # 作成日を参照し、期間が経過していたらcsv and img フォルダ削除
    add_functions.delete_old_files(specified_date)

    # Use browser cache ======================================================
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-browser-side-navigation")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-site-isolation-trials")
    options.add_argument("--disable-hang-monitor")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-backgrounding-occluded-windows")
    # ======================================================================
    # use headress browser (ヘッドレスブラウザは全ジャンル、ページ取得するまで開いたままにする・時間短縮)
    options.add_argument('--headless')  # ヘッドレスを解除する場合はコメントアウト

    # ブラウザ起動
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    # Wait for page source to load ※ソースを読み込むまで待機
    driver.implicitly_wait(15)

    # 設定ファイルからジャンル、ジャンルID読み込み～スクレイピング～CSV作成
    for row in csv.reader(open(genre_file, 'r', encoding='utf-8_sig', newline='')):
        print(f'genre:{row[0]}   genre_id:{row[1]}')
        # スクレイピングとCSV書き込み
        csv_save(row[0], row[1], intervaltime, driver)

    driver.close()  # ブラウザを閉じる

    # エクセルへの書き込み実行=================================================
    # 【重要】マクロを実行して写真を貼り付ける場合、同じディレクトリにimgフォルダがないとエクセルの画像が
    #   消失してしまうので、画像ファイルを先にoutput+timespan+日付+imgフォルダに移動させ、
    #   テンプレートのエクセルファイルを移動させてからマクロを実行する。
    #   全部の処理後にリネームしないと、マクロがエラーを起こす

    # 処理後のエクセルファイルの保存先ディレクトリ　例）output/realtime/日付フォルダ
    output_ex_files_dir = os.path.join(output_dir, intervaltime, add_datetime())

    # 写真ファイルのコピー dat/imgからoutpu/intervaltime/datetime/img
    print('copy img files')
    shutil.copytree(img_dir, os.path.join(output_ex_files_dir, 'img'))

    print('==========エクセル処理開始==========')
    time.sleep(2)
    # CSVからエクセルへ書き込み＆マクロ実行
    export_ex(output_ex_files_dir, intervaltime)

    # 処理時間result
    print('\n==========全処理終了==========')
    end_time = add_datetime()
    print(f'開始時間：{start_time}')
    print(f'終了時間：{end_time}')

# ===========================================================================

if __name__ == '__main__':

    # テスト、本番選択（テスト用はジャンルが３種類,リアルタイム実行）
    # 1はテスト、２は本番
    mode_b = 2

    # time_table import
    time_list = []
    [time_list.append(row)
     for row in csv.reader(open(time_table_file, 'r', encoding='utf-8_sig', newline=''))]

    if mode_b == 1:
        # ===for test===
        while True:
            main_func(mode=1, mode2=mode_b)
            print('\n*****待機中*****\n')
            time.sleep(180)

    elif mode_b == 2:
        # timer
        print(time_list)
        # realtime
        [schedule.every().day.at(t).do(main_func, mode=1, mode2=mode_b) for t in time_list[0][1:]]
        # daily
        [schedule.every().day.at(t2).do(main_func, mode=2, mode2=mode_b) for t2 in time_list[1][1:]]
        # weekly
        [schedule.every().monday.at(t3).do(main_func, mode=3, mode2=mode_b) for t3 in time_list[2][1:]]

        print('\n==== ratenk start====')
        print('==== 処理時間まで待機==')

        while True:
            schedule.run_pending()
            time.sleep(1)
    else:
        print('デバッグ指定エラー')
