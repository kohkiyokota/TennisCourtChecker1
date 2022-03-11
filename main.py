from curses.ascii import NUL
from pty import slave_open
from xml.etree.ElementPath import xpath_tokenizer
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome import service

import config
import time
from datetime import datetime, date, timedelta
import jpholiday

# from ast import Or
# from urllib import request  # urllib.requestモジュールをインポート

# 自作モジュール
from modules.sendLine import send_line_notify


# # ################################
# # スプレッドシート読み込み
# # ################################
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope =['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

spreadsheet = client.open('TennisCourtChecker') # 操作したいスプレッドシートの名前を指定する
worksheet = spreadsheet.worksheet('東京都スポーツ施設サービス') # シートを指定する

# Extract and print all of the values
# list_of_hashes = sheet.get_all_records()
# print(list_of_hashes)

# Headless Chromeをあらゆる環境で起動させるオプション
# 省メモリ化しないとメモリ不足でクラッシュする
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--remote-debugging-port=9222')
# options.add_argument('window-size=500,500')
# UA指定しないとTikTok弾かれる
UA = 'SetEnvIfNoCase User-Agent "NaverBot" getout'
options.add_argument(f'user-agent={UA}')

result = []



def writeSheet(data):
    worksheet.delete_rows(2) # 2行目(計測時間)を削除
    worksheet.delete_rows(1) # 1行目(空き状況)を削除
    worksheet.append_row(data) # dataを最終行に挿入

def main():
    # 処理時間計測①：開始
    start_time = time.perf_counter()

    for i in range(3):  # 最大3回実行
        try:
            # 「東京都スポーツ施設サービス」のURL
            url1 = 'https://yoyaku.sports.metro.tokyo.lg.jp/user/view/user/homeIndex.html'

            # # ################################
            # # 日時を取得・計算
            # # ################################
            today = datetime.today()
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            today_year =  int('20' + datetime.strftime(today, '%y'))

            today_month = datetime.strftime(today, '%m')
            if today_month[0] == '0':
              today_month = today_month[1:2]
            today_month = int(today_month)

            today_day = datetime.strftime(today, '%d')
            if today_day[0] == '0':
              today_day = today_day[1:2]
            today_day = int(today_day)

            if today_month == 12:
              next_month = 1
              next_year = today_year + 1
            else:
              next_month = today_month + 1 # 来月
              next_year = today_year

            # today_year
            # today_month
            # today_day
            # next_month
            # next_year

            # # ################################
            # # 設定をもとに曜日と日時の配列を作成
            # # ################################

            # 曜日と日時の配列
            dayOfWeek_array = calcDayOfWeek()
            print(dayOfWeek_array)

            this_holidays = []
            next_holidays = []

            # もし曜日指定に「祝日」が入っていたら
            if len(config.HOL) > 0:
                print('祝日入ってる')
                # 今月の祝日を取得
                for hl1 in jpholiday.month_holidays(today_year, today_month):
                  this_holidays.append(hl1[0].day)

                # 来月の祝日を取得 
                for hl2 in jpholiday.month_holidays(next_year, next_month):
                  next_holidays.append(hl2[0].day)

            print(this_holidays)
            print(next_holidays)

            # # ################################
            # # 公園名ごとにデータを取得（下準備）
            # # ################################

            # 環境変数に設定されている公園を取得（東京都スポーツ施設サービス）
            PARKS = config.PARKS
            parks_array = PARKS.split(',')

            driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options=options)
            driver.implicitly_wait(60)
            driver.get(url1)

            # step1: 公園選択画面へ行く
            btn = driver.find_element(by=By.XPATH, value='//*[@id="nameSearch"]')
            btn.click()

            # # ################################
            # # 公園名ごとにデータを取得
            # # ################################
            for park in parks_array:
                parks_nums = len(driver.find_elements(by=By.XPATH, value='//*[@id="resultItems"]/tr'))
                for i in range(parks_nums):
                    # ①合致する公園名があれば次へ
                    if driver.find_element(by=By.XPATH, value=f'//*[@id="resultItems"]/tr[{str(i + 1)}]/td[1]/span').text == park:
                        print(f'「{park}」の情報の取得を開始')
                        # result.append(park)
                        # 公園のページへ移動
                        btn_list = driver.find_elements(by=By.XPATH, value='//*[@id="srchBtn"]')                
                        btn_list[i].click()

                        # カレンダーの段数を計算
                        rows_num = len(driver.find_elements(by=By.XPATH, value='//*[@id="calendar"]/table[2]/tbody/tr'))

                        for row in range(rows_num):
                            # １段目は曜日なので飛ばす
                            if row > 0:
                                # 指定された曜日を検索
                                for dow_se in dayOfWeek_array:
                                    # 指定された曜日であれば
                                    if dow_se != 0:
                                        day = driver.find_element(by=By.XPATH, value=f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{dayOfWeek_array.index(dow_se) + 1}]')
                                        # 日付があり、今日以降で、祝日じゃなければ日付をクリックして空きをチェック
                                        if day.text and int(day.text) >= today_day and int(day.text) not in this_holidays:
                                            date = day.text # 日付：クリックすると変わってしまうから変数に入れとく
                                            day.click()
                                            # 空きをチェック
                                            checkEmpty(driver, park, today_month, date, dow_se)

                        # 祝日を検索
                        for h in this_holidays:
                            if h >= today_day:
                                for row in range(rows_num):
                                    if row > 0:
                                        for col in range(6):
                                            d = driver.find_element(by=By.XPATH, value=f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{col + 1}]')
                                            if len(d.text) > 0 and int(d.text) == h:
                                                d.click()
                                                # 空きをチェック
                                                hol_se = config.HOL.split('-')
                                                if col == 0:
                                                    hol_se.insert(0, '日祝')
                                                if col == 1:
                                                    hol_se.insert(0, '月祝')
                                                if col == 2:
                                                    hol_se.insert(0, '火祝')
                                                if col == 3:
                                                    hol_se.insert(0, '水祝')
                                                if col == 4:
                                                    hol_se.insert(0, '木祝')
                                                if col == 5:
                                                    hol_se.insert(0, '金祝')
                                                if col == 6:
                                                    hol_se.insert(0, '土祝')
                                                checkEmpty(driver, park, today_month, date, hol_se)

                        # 翌月へ移動
                        print('今月最終日なので翌月に移動します')
                        next_btn = driver.find_element(by=By.XPATH, value='//*[@id="calendar"]/table[1]/tbody/tr/td/div/a')
                        next_btn.click()
                        print('翌月のチェックを開始')
                        checkNextMonth(driver, dayOfWeek_array, park, next_month, next_holidays)

                        # 翌月のチェック終わったら公園名の選択画面へ戻る
                        driver.get(url1)
                        btn = driver.find_element(by=By.XPATH, value='//*[@id="nameSearch"]')
                        btn.click()

            # 結果をフォーマット
            result.sort()
            date1 = ''
            final_result = []
            for item in result:
                date2 = item.split('_')[0]
                others = item.split('_')[1]
                if date1 == date2:
                    final_result.append(others)
                else:
                    final_result.append(date2)
                    final_result.append(others)
                    date1 = date2

            # spreadsheetの1業目を取得（これとfinal resultを比較！）
            history = worksheet.row_values(1)
            print('取得結果')
            print(final_result)
            print('前回の取得結果')
            print(history)

            # 変更チェック
            if history == final_result:
                print('前回と変更なし')
            # 空きが出た場合
            elif len(history) <= len(final_result):
                print('コート増えた：通知あり')
                writeSheet(final_result)
                # LINE通知
                message = '\n【テニスコート空き状況：都営】\n'
                for item2 in final_result:
                    #日付なら改行入れる
                    if '〜' not in item2:
                        message += '\n'
                    message += f'{item2}\n'
                send_line_notify(message)
            # コートに空きがない場合
            elif len(final_result) == 0:
                print('空きなし：通知あり')
                # worksheet.delete_rows(0) # 1行目(空き状況)を削除
                writeSheet(final_result)
                send_line_notify('空きコートはありません。')
            # コートが埋まった場合（通知なし）
            else:
                print('コートへった：通知なし')
                writeSheet(final_result)

            # 処理時間計測②：修了
            end_time = time.perf_counter()
            # 経過時間を出力(秒)
            elapsed_time = end_time - start_time
            worksheet.update_acell('A2', now)
            worksheet.update_acell('B2', elapsed_time)
            print(elapsed_time)
            # worksheet.append_row() # 時間を最終行(2行目)に挿入

        except Exception as e:
            # import traceback
            # traceback.print_exc()
            if i == 2:
                err_title = e.__class__.__name__ # エラータイトル
                message = f'例外発生！\n\n{err_title}\n{e.args}'
                send_line_notify(message, config.LNT_FOR_ERROR)
            pass

        else: # 例外が発生しなかった時だけ実行される
            break  # 失敗しなかった時はループを抜ける

# 翌月のチェック
def checkNextMonth(driver, dayOfWeek_array, park, month, next_holidays):
    # カレンダーが何段か計算
    rows_num2 = len(driver.find_elements(by=By.XPATH, value='//*[@id="calendar"]/table[2]/tbody/tr'))

    for row in range(rows_num2):
        # １段目は曜日なので飛ばす
        if row > 0:
            for dow_se in dayOfWeek_array:
                # 指定された曜日であれば
                if dow_se != 0:
                    day = driver.find_element(by=By.XPATH, value=f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{dayOfWeek_array.index(dow_se) + 1}]')
                    # 日付があ理、祝日じゃなければ日付をクリックして空きをチェック
                    if day.text and int(day.text) not in next_holidays:
                        date2 = day.text # 日付：クリックすると変わってしまうから変数に入れとく
                        day.click()
                        # 空きをチェック
                        checkEmpty(driver, park, month, date2, dow_se)

    # 祝日を検索
    for h in next_holidays:
            for row in range(rows_num2):
                if row > 0:
                    for col in range(6):
                        d = driver.find_element(by=By.XPATH, value=f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{col + 1}]')
                        if len(d.text) > 0 and int(d.text) == h:
                            d.click()
                            # 空きをチェック
                            hol_se = config.HOL.split('-')
                            if col == 0:
                                hol_se.insert(0, '日祝')
                            if col == 1:
                                hol_se.insert(0, '月祝')
                            if col == 2:
                                hol_se.insert(0, '火祝')
                            if col == 3:
                                hol_se.insert(0, '水祝')
                            if col == 4:
                                hol_se.insert(0, '木祝')
                            if col == 5:
                                hol_se.insert(0, '金祝')
                            if col == 6:
                                hol_se.insert(0, '土祝')
                            checkEmpty(driver, park, month, date, hol_se)

    print('翌月のチェック終了')

# 設定をもとに曜日と時間帯の配列を作成
def calcDayOfWeek():
    dayOfWeek_array = []
    if len(config.SUN) > 0:
        dayOfWeek_array.append(config.SUN.split('-'))
        dayOfWeek_array[0].insert(0, '日')
    else:
        dayOfWeek_array.append(NUL)
    if len(config.MON) > 0:
        dayOfWeek_array.append(config.MON.split('-'))
        dayOfWeek_array[1].insert(0, '月')
    else:
        dayOfWeek_array.append(NUL)
    if len(config.TUE) > 0:
        dayOfWeek_array.append(config.TUE.split('-'))
        dayOfWeek_array[2].insert(0, '火')
    else:
        dayOfWeek_array.append(NUL)
    if len(config.WED) > 0:
        dayOfWeek_array.append(config.WED.split('-'))
        dayOfWeek_array[3].insert(0, '水')
    else:
        dayOfWeek_array.append(NUL)
    if len(config.THU) > 0:
        dayOfWeek_array.append(config.THU.split('-'))
        dayOfWeek_array[4].insert(0, '木')
    else:
        dayOfWeek_array.append(NUL)
    if len(config.FRI) > 0:
        dayOfWeek_array.append(config.FRI.split('-'))
        dayOfWeek_array[5].insert(0, '金')
    else:
        dayOfWeek_array.append(NUL)
    if len(config.SAT) > 0:
        dayOfWeek_array.append(config.SAT.split('-'))
        dayOfWeek_array[6].insert(0, '土')
    else:
        dayOfWeek_array.append(NUL)
    return dayOfWeek_array

# 空き状況を検索
def checkEmpty(driver, park, month, day, se):
    syumoku_list = driver.find_elements(by=By.XPATH, value='//*[@id="ppsname"]')
    # 種目にテニスがあるかチェック
    for syumoku in syumoku_list:
        START = int(se[1])
        END = int(se[2])

        # ②テニス（人工芝）があれば次へ
        if syumoku.text == 'テニス（人工芝）':
            syumoku_index = syumoku_list.index(syumoku)
            times_num = len(driver.find_elements(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td'))
            for time_index in range(times_num):
                td = int(driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td[{time_index + 1}]').text)
                # ③もし0以上のセルがあれば次へ
                if td > 0:
                    # 空いてる時間帯
                    t = int(driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]').text.replace(':00', ''))
                    # ④もし設定した時間内なら通知
                    if t >= START and t <= END:
                        park_name = park # 公園名
                        month_day = f'{month}/{day}（{se[0]}）'
                        # 開始時刻
                        start = driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]').text 
                        # 終了時刻
                        if time_index + 1 == times_num:
                          before_start = int(driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index}]').text.replace(':00', ''))
                          end = str(int(start.replace(':00', '')) + (int(start.replace(':00', '')) - before_start)) + ':00'
                        else:
                            end = driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 2}]').text 
                        result.append(f'{month_day}_{start}〜{end}@{park} {td}面')

        # ②テニス（ハード）があれば次へ
        if syumoku.text == 'テニス（ハード）':
            syumoku_index = syumoku_list.index(syumoku)
            times_num = len(driver.find_elements(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td'))
            for time_index in range(times_num):
                td = int(driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td[{time_index + 1}]').text)
                # ③もし0以上のセルがあれば次へ
                if td > 0:
                    # 空いてる時間帯
                    t = int(driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]').text.replace(':00', ''))
                    # ④もし設定した時間内なら通知
                    if t >= START and t <= END:
                        park_name = park # 公園名
                        month_day = f'{month}/{day}（{se[0]}）'
                        # 開始時刻
                        start = driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]').text 
                        # 終了時刻
                        if time_index + 1 == times_num:
                          before_start = int(driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index}]').text.replace(':00', ''))
                          end = str(int(start.replace(':00', '')) + (int(start.replace(':00', '')) - before_start)) + ':00'
                        else:
                            end = driver.find_element(by=By.XPATH, value=f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 2}]').text 
                        result.append(f'{month_day}_{start}〜{end}@{park} {td}面')


if __name__ == "__main__":
    main()

