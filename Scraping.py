#!/usr/bin/env python
# coding: utf-8

# In[21]:

import subprocess
import os
import time
import datetime as dt
import smtplib
import pyautogui as py #キーボード操作で有名なライブラリです。
import pyperclip       #Pythonでクリップボードを操作する（コピー＆ペースト）時に使用するライブラリ
import requests
from bs4 import BeautifulSoup as bs
import re
from email.mime.text import MIMEText
from email.utils import formatdate

tdatetime = dt.datetime.now()
tstr = tdatetime.strftime('%Y/%m/%d')

f = open('./Secret.txt','r')                                   #パスワードを取得する
datalist = f.readlines()
f.close()

#ログ出力開始
f = open('./log.txt','w', encoding='UTF-8')
f.write (tstr + "実行分作業、はじまるドン！！")
f.close()

PASSWORD = datalist[1]
FROM_ADDRESS = datalist[3]
TO_ADDRESS = datalist[5]
APP_PASSWORD = datalist[7]

# In[47]:


def create_message(from_addr, to_addr, subject, body):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Date'] = formatdate()
    return msg


# In[48]:


def send_mail(from_addr, to_addr, body_msg):
    smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpobj.ehlo()
    smtpobj.starttls()
    smtpobj.ehlo()
    smtpobj.login(FROM_ADDRESS, APP_PASSWORD)
    smtpobj.sendmail(from_addr, to_addr, body_msg.as_string())
    smtpobj.close()


# In[44]:


def start_market_speed():
    py.hotkey("win","m")  #全ウィンドウの最小化、Win + m の操作と同じ
    os.chdir(r'C:\Users\XXXXXX\AppData\Local\MarketSpeed2\Download')         #pyファイルのあるディレクトリへ移動してから、exeファイルを起動する
    global market_speed
    #exeファイルのパスを下記に格納
    market_speed_path = r'C:\Users\XXXXXX\AppData\Local\MarketSpeed2\Bin\MarketSpeed2.exe'
    market_speed = subprocess.Popen(market_speed_path)
    time.sleep(5)
    py.click(1280,800)
    time.sleep(1)
    #ログインID/PWを設定
    py.typewrite(PASSWORD)

    time.sleep(2)
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("Enter")
    time.sleep(10)

    py.hotkey("win","up")
    time.sleep(2)

    py.click(76,778)
    time.sleep(2)
    py.click(106,529)
    time.sleep(2)
    py.click(340,227)
    time.sleep(8)

    #---------[同意する]ボタンの選択---------
    py.press("tab")
    time.sleep(1)
    py.hotkey("enter")
    time.sleep(3)
    #---------[同意する]ボタンの選択---------

    #---------[今日の新聞]タブの選択---------
    for i in range(8):            #[日本経済新聞朝刊]まで
	    py.press("tab")
    time.sleep(1)
    py.hotkey("enter")
    #---------[今日の新聞]タブの選択---------

    #---------ページ全体のテキストを取得---------
    py.hotkey("ctrl","a")
    py.hotkey("ctrl","c")
    text = pyperclip.paste()
   #---------ページ全体のテキストを取得---------

    #---------[条件をクリア]タブまで移動---------
    for i in range(12):            #[日本経済新聞朝刊]まで
        py.press("tab")
    time.sleep(1)

    text = re.sub('\n', ' ', text)
    dates = re.search(r'日付(.+)媒体', text).group(0)   #日付欄
    dates_list = dates.split()
    print(dates_list.pop(0))
    print(dates_list.pop(-1))
    for date in dates_list:
        print(date)
        py.press("tab")
    time.sleep(1)
    
    medium = re.search(r'媒体(.+)キーワードを入力してください', text).group(0)   #媒体欄
    medium_list = medium.split()
    print(medium_list.pop(0))
    print(medium_list.pop(-1))
    for medium in medium_list:
        print(medium)
        py.press("tab")

    for i in range(2):            #[条件をクリア]まで
        py.press("tab")
    #---------[条件をクリア]タブまで移動---------

    #---------[今日の新聞]ページのトピック数を取得---------
    time.sleep(1)
    topicks = re.search(r'全ての記事(.+)文化\(', text).group(0)
    topicks_list = topicks.split()
    #---------[今日の新聞]ページのトピック数を取得---------

    #---------[一面]チェックボックスまで移動---------
    for topicks in topicks_list:
        py.press("tab")
    py.press("space")
    #---------[一面]チェックボックスまで移動---------

    #---------[本文を表示]ボタンの選択---------
    time.sleep(1)
    py.hotkey("shift","tab")
    time.sleep(1)
    py.press("space")
    #---------[本文を表示]ボタンの選択---------

    time.sleep(1)

    #---------テキスト全体を選択---------
    time.sleep(1)
    py.hotkey("ctrl","a")
    py.hotkey("ctrl","c")
    #---------テキスト全体を選択---------

    #---------余分な文言を削除---------
    text = pyperclip.paste()
    edited_text = text.replace('日経テレコン２１ヘルプとサポートログアウト','').replace('記事検索','').replace('検索詳細条件を指定','').replace('ニュース','').replace('きょうの新聞','').replace('記事検索','').replace('キーワードを入力してください','').replace('印刷','').replace('PDF','').replace('一覧に戻る','').replace('ご提供する情報について 個人情報の取り扱いについて ','').replace('本サービスに関する知的所有権その他一切の権利は日本経済新聞社またはその情報提供者に帰属します。','').replace('また本サービスは方法の如何、有償無償を問わず契約者以外の第三者に利用させることはできません。','').replace('Copyright © 日本経済新聞社 Nikkei Inc. All Rights Reserved.','').replace('','')
    #---------余分な文言を削除---------

    return edited_text

#本体処理
body = start_market_speed()
subject = "朝刊一面" + tstr
body_msg = create_message(FROM_ADDRESS, TO_ADDRESS, subject, body)
send_mail(FROM_ADDRESS, TO_ADDRESS, body_msg)
