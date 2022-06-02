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
import os.path
import time
from datetime import datetime, date, timedelta
import jpholiday
import calendar

# 重複防止用
import fcntl

# 自作モジュール
from modules.sendLine import send_line_notify

# # ################################
# # スプレッドシートから設定を読み込む
# # ################################
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope =['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
# 認証情報を相対パスで取得・設定
# dirname = os.getcwd()
# cred_dir = os.path.join(dirname, 'client_secret.json')
creds = ServiceAccountCredentials.from_json_keyfile_name('/home/admin/TennisCourtChecker/client_secret.json', scope)
# creds = ServiceAccountCredentials.from_json_keyfile_name(cred_dir, scope)
client = gspread.authorize(creds)

spreadsheet = client.open('TennisCourtChecker') # 操作したいスプレッドシートの名前を指定する
worksheet = spreadsheet.worksheet('東京都スポーツ施設サービス') # シートを指定する

configSheet = spreadsheet.worksheet('設定_都営')
SUN = configSheet.acell('C3').value
MON = configSheet.acell('C4').value
TUE = configSheet.acell('C5').value
WED = configSheet.acell('C6').value
THU = configSheet.acell('C7').value
FRI = configSheet.acell('C8').value
SAT = configSheet.acell('C9').value
HOL = configSheet.acell('C10').value
TOKEN = configSheet.acell('C14').value
LNT_FOR_ERROR = configSheet.acell('C15').value

result = []

# ###########################
# スプレッドシートに記載する
# ###########################
def writeSheet(data):
    worksheet.delete_rows(2) # 2行目(計測時間)を削除
    worksheet.delete_rows(1) # 1行目(空き状況)を削除
    worksheet.append_row(data) # dataを最終行に挿入

# ###########################
# 要素を取得する（単数）
# ###########################
def getElement(xpath, wait_second, retries_count):
    error = ''
    for _ in range(retries_count):
        try:
            # 失敗しそうな処理
            selector = xpath
            element = WebDriverWait(driver, wait_second).until(
              EC.visibility_of_element_located((By.XPATH, selector))
            )
        except Exception as e:
            # エラーメッセージを格納する
            error = e
        else:
            # 失敗しなかった場合は、ループを抜ける
            break
    else:
        # リトライが全部失敗したときの処理。エラー内容(error)や実行時間、操作中のURL、セレクタ、スクショなどを通知する。
        send_line_notify(error, LNT_FOR_ERROR)
        exit() # プログラムを強制終了する
    return element

# ###########################
# 要素を取得する（複数）
# ###########################
def getElements(xpath, wait_second, retries_count):
    error = ''
    for _ in range(retries_count):
        try:
            # 失敗しそうな処理
            selector = xpath
            elements = WebDriverWait(driver, wait_second).until(
              EC.visibility_of_all_elements_located((By.XPATH, selector))
            )
        except Exception as e:
            # エラーメッセージを格納する
            error = e
        else:
            # 失敗しなかった場合は、ループを抜ける
            break
    else:
        # リトライが全部失敗したときの処理。エラー内容(error)や実行時間、操作中のURL、セレクタ、スクショなどを通知する。
        send_line_notify(error, LNT_FOR_ERROR)
        exit() # プログラムを強制終了する
    return elements

def main():
    # 処理時間計測①：開始
    start_time = time.perf_counter()

    for x in range(3):  # 最大3回実行
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
            # # 22日までなら当月、それ以降なら翌月
            # # ################################
            if today_day < 22:
                print('今月まで')
                goNextMonth = False
            else:
                print('来月まで')
                goNextMonth = True

            # # ################################
            # # 設定をもとに曜日と日時の配列を作成
            # # ################################

            # 曜日と日時の配列
            dayOfWeek_array = calcDayOfWeek()
            print(dayOfWeek_array)

            this_holidays = []
            next_holidays = []


            # もし曜日指定に「祝日」が入っていたら
            if len(HOL) > 1:
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
            PARKS = configSheet.acell('C12').value
            parks_array = PARKS.split(',')

            driver.get(url1)

            time.sleep(5)

            # # ################################
            # # 公園名ごとにデータを取得
            # # ################################
            for park in parks_array:

                # step1: 公園選択画面へ行く
                btn = getElement('//*[@id="nameSearch"]', 20, 3)
                btn.click()

                parks_nums = len(getElements('//*[@id="resultItems"]/tr', 20, 3))
                for i in range(parks_nums):
                    # ①合致する公園名があれば次へ
                    if getElement(f'//*[@id="resultItems"]/tr[{str(i + 1)}]/td[1]/span', 20, 3).text == park:
                        print(f'「{park}」の情報の取得を開始')
                        # 公園のページへ移動
                        btn_list = getElements('//*[@id="srchBtn"]', 20, 3)
                        btn_list[i].click()

                        # カレンダーの段数を計算
                        rows_num = len(getElements('//*[@id="calendar"]/table[2]/tbody/tr', 20, 3))

                        for row in range(rows_num):
                            # １段目は曜日なので飛ばす
                            if row > 0:
                                # 指定された曜日を検索
                                for dow_se in dayOfWeek_array:
                                    # 指定された曜日であれば
                                    if dow_se != 0:
                                        day = getElement(f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{dayOfWeek_array.index(dow_se) + 1}]', 20, 3)
                                        # 日付があり、今日以降で、祝日じゃなければ日付をクリックして空きをチェック
                                        if day.text and int(day.text) >= today_day and int(day.text) not in this_holidays:
                                            date = day.text # 日付：クリックすると変わってしまうから変数に入れとく
                                            day.click()
                                            # 空きをチェック
                                            checkEmpty(park, today_month, int(date), dow_se)

                        # 祝日を検索
                        for h in this_holidays:
                            if h >= today_day:
                                for row in range(rows_num):
                                    if row > 0:
                                        for col in range(6):
                                            d = getElement(f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{col + 1}]', 20, 3)
                                            if len(d.text) > 0 and int(d.text) == h:
                                                date = d.text # 日付：クリックすると変わってしまうから変数に入れとく
                                                d.click()
                                                # 空きをチェック
                                                hol_se = HOL.split('-')
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
                                                checkEmpty(park, today_month, int(date), hol_se)

                        if goNextMonth:
                            # 翌月へ移動
                            print('今月最終日なので翌月に移動します')
                            next_btn = getElement('//*[@id="calendar"]/table[1]/tbody/tr/td/div/a', 20, 3)
                            next_btn.click()
                            print('翌月のチェックを開始')
                            checkNextMonth(dayOfWeek_array, park, next_month, next_holidays)

                        # 現在の公園のチェック終わったら公園名の選択画面へ戻る
                        driver.get(url1)
                        btn2 = getElement('//*[@id="nameSearch"]', 20, 3)
                        btn2.click()

                        time.sleep(5)


            # Chromeを終了
            # driver.close()
            # driver.quit()

            # 結果をフォーマット
            result.sort()
            date1 = ''
            final_result = []
            for item in result:
                date2 = item.split('_')[0]
                others = item.split('_')[1]
                if others[0] == '0':
                    others = others[1:]
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

                # #######################################
                # LINE通知 ###############################
                # #######################################
                result_length = len(final_result) # 結果数
                pages = result_length // 40 + 1    # LINEの通数

                for page in range(pages):
                    if page == 0:
                        message = '\n【テニスコート空き状況：都営】\n'
                    else:
                        message = '\n\n'

                    m = final_result[40 * page: 40 * (page + 1)]
                    for item2 in m:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message += '\n'
                        message += f'{item2}\n'
                    
                    send_line_notify(message, TOKEN)
                    send_line_notify(message, LNT_FOR_ERROR)

            # コートに空きがない場合
            elif len(final_result) == 0:
                print('空きなし：通知あり')
                # worksheet.delete_rows(0) # 1行目(空き状況)を削除
                writeSheet(final_result)
                send_line_notify('空きコートはありません。', TOKEN)
                send_line_notify('空きコートはありません。', LNT_FOR_ERROR)
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
            # driver.close()
            # driver.quit()
            if x == 2:
                err_title = e.__class__.__name__ # エラータイトル
                message = f'\n【都営】\n\n例外発生！\n\n{err_title}\n{e.args}'
                send_line_notify(message, LNT_FOR_ERROR)
            pass

        else: # 例外が発生しなかった時だけ実行される
            break  # 失敗しなかった時はループを抜ける

# 翌月のチェック
def checkNextMonth(dayOfWeek_array, park, month, next_holidays):
    # カレンダーが何段か計算
    rows_num2 = len(getElements('//*[@id="calendar"]/table[2]/tbody/tr', 20, 3))

    for row in range(rows_num2):
        # １段目は曜日なので飛ばす
        if row > 0:
            for dow_se in dayOfWeek_array:
                # 指定された曜日であれば
                if dow_se != 0:
                    day = getElement(f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{dayOfWeek_array.index(dow_se) + 1}]', 20, 3)
                    # 日付があ理、祝日じゃなければ日付をクリックして空きをチェック
                    if day.text and int(day.text) not in next_holidays:
                        date2 = day.text # 日付：クリックすると変わってしまうから変数に入れとく
                        day.click()
                        # 空きをチェック
                        checkEmpty(park, month, int(date2), dow_se)

    # 祝日を検索
    for h in next_holidays:
            for row in range(rows_num2):
                if row > 0:
                    for col in range(6):
                        d = getElement(f'//*[@id="calendar"]/table[2]/tbody/tr[{row + 1}]/td[{col + 1}]', 20, 3)
                        if len(d.text) > 0 and int(d.text) == h:
                            date3 = d.text
                            d.click()
                            # 空きをチェック
                            hol_se = HOL.split('-')
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
                            checkEmpty(park, month, int(date3), hol_se)

    print('翌月のチェック終了')

# 設定をもとに曜日と時間帯の配列を作成
def calcDayOfWeek():
    dayOfWeek_array = []
    if len(SUN) > 1:
        dayOfWeek_array.append(SUN.split('-'))
        dayOfWeek_array[0].insert(0, '日')
    else:
        dayOfWeek_array.append(NUL)
    if len(MON) > 1:
        dayOfWeek_array.append(MON.split('-'))
        dayOfWeek_array[1].insert(0, '月')
    else:
        dayOfWeek_array.append(NUL)
    if len(TUE) > 1:
        dayOfWeek_array.append(TUE.split('-'))
        dayOfWeek_array[2].insert(0, '火')
    else:
        dayOfWeek_array.append(NUL)
    if len(WED) > 1:
        dayOfWeek_array.append(WED.split('-'))
        dayOfWeek_array[3].insert(0, '水')
    else:
        dayOfWeek_array.append(NUL)
    if len(THU) > 1:
        dayOfWeek_array.append(THU.split('-'))
        dayOfWeek_array[4].insert(0, '木')
    else:
        dayOfWeek_array.append(NUL)
    if len(FRI) > 1:
        dayOfWeek_array.append(FRI.split('-'))
        dayOfWeek_array[5].insert(0, '金')
    else:
        dayOfWeek_array.append(NUL)
    if len(SAT) > 1:
        dayOfWeek_array.append(SAT.split('-'))
        dayOfWeek_array[6].insert(0, '土')
    else:
        dayOfWeek_array.append(NUL)
    return dayOfWeek_array

# 空き状況を検索
def checkEmpty(park, month, day, se):
    syumoku_list = getElements('//*[@id="ppsname"]', 20, 3)

    # ソート用に月と日を2桁にする
    if int(month) < 10:
        month = '0' + str(month)
    if int(day) < 10:
        day = '0' + str(day)

    # 種目にテニスがあるかチェック
    for syumoku in syumoku_list:
        START = int(se[1])
        END = int(se[2])

        # ②テニス（人工芝）があれば次へ
        if syumoku.text == 'テニス（人工芝）':
            syumoku_index = syumoku_list.index(syumoku)
            times_num = len(getElements(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td', 20, 3))
            for time_index in range(times_num):
                td = int(getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td[{time_index + 1}]', 20, 3).text)
                # ③もし0以上のセルがあれば次へ
                if td > 0:
                    # 空いてる時間帯
                    t = int(getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]', 20, 3).text.replace(':00', ''))
                    # ④もし設定した時間内なら通知
                    if t >= START and t <= END:
                        park_name = park # 公園名

                        month_day = f'{month}/{day}（{se[0]}）'
                        # 開始時刻
                        start =  getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]', 20, 3).text
                        # 終了時刻
                        if time_index + 1 == times_num:
                          before_start = int(getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index}]', 20, 3).text.replace(':00', ''))
                          end = str(int(start.replace(':00', '')) + (int(start.replace(':00', '')) - before_start)) + ':00'
                        else:
                            end = getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 2}]', 20, 3).text 
                        if start == '9:00':
                            start = '09:00'
                        result.append(f'{month_day}_{start}〜{end}@{park} {td}面')

        # ②テニス（ハード）があれば次へ
        if syumoku.text == 'テニス（ハード）':
            syumoku_index = syumoku_list.index(syumoku)
            times_num = len(getElements(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td', 20, 3))
            for time_index in range(times_num):
                td = int(getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[3]/td[{time_index + 1}]', 20, 3).text)
                # ③もし0以上のセルがあれば次へ
                if td > 0:
                    # 空いてる時間帯
                    t = int(getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]', 20, 3).text.replace(':00', ''))
                    # ④もし設定した時間内なら通知
                    if t >= START and t <= END:
                        park_name = park # 公園名

                        month_day = f'{month}/{day}（{se[0]}）'
                        # 開始時刻
                        start =  getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 1}]', 20, 3).text
                        # 終了時刻
                        if time_index + 1 == times_num:
                          before_start = int(getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index}]', 20, 3).text.replace(':00', ''))
                          end = str(int(start.replace(':00', '')) + (int(start.replace(':00', '')) - before_start)) + ':00'
                        else:
                            end = getElement(f'//*[@id="isNotEmptyPager"]/table[{syumoku_index + 1}]/tbody/tr[2]/td[{time_index + 2}]', 20, 3).text 
                        if start == '9:00':
                            start = '09:00'
                        result.append(f'{month_day}_{start}〜{end}@{park} {td}面')


if __name__ == "__main__":
    lockfilePath = 'lockfile.lock'
    with open(lockfilePath , "w") as lockFile:
        try:
            fcntl.flock(lockFile, fcntl.LOCK_EX | fcntl.LOCK_NB)
            # Do SOMETHING
            # Headless Chromeをあらゆる環境で起動させるオプション
            # 省メモリ化しないとメモリ不足でクラッシュする
            options = Options()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            # options.add_argument('--remote-debugging-port=9222')
            # さくらでchrome not reachableエラー出たから追加
            # options.add_argument("--single-process") 
            # options.add_argument("--disable-setuid-sandbox") 
            # options.add_argument('window-size=500,500')
            # UA指定しないとTikTok弾かれる
            UA = 'SetEnvIfNoCase User-Agent "NaverBot" getout'
            options.add_argument(f'user-agent={UA}')
            chrome_service = service.Service(executable_path=ChromeDriverManager().install())
            driver = webdriver.Chrome(service=chrome_service, options=options) 
            driver.implicitly_wait(20)
            main()
            driver.delete_all_cookies()
            driver.close()
            driver.quit()

        except IOError:
            print('process already exists')
