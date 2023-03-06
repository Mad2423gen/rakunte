# rakunte
楽天ランキングスクレイピング
APIを使用するとリアルタイムのデータしか取得できないのでseleniumを使用してデイリーとウィークリーを取得する
出力はcsvとエクセルファイル
動作環境：Windows11 or Window10 Python3.10　GoogleChrome
モジュールは　pip install -r requirements.txt でインストール　（venv等、仮想環境使用推奨）

使用方法
1.timetable.csvに巡回時刻を設定　４行目のupdate_interval_day　はcsvの初期化タイミング。それまでは差分を保存していく
2.main2.4.pyの386行目、mode_b を設定し、コマンドプロンプトから py main2.4.py で実行
実行結果は rakunte/output/rialtime or dayly or weekly/datetimeフォルダに出力される。
エクセルの写真が消えるので、同梱されているimgフォルダは削除しないこと。

除外キーワードについて
rakunte/config/keywordディレクトリが生成されるので、ジャンルごとに除外キーワードを一行ずつ入れておく


