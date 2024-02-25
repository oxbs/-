#必要ライブラリのインポート
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from time import sleep,time
import pandas as pd
from bs4 import BeautifulSoup
import requests
import re
import math

#chromedriverの呼出し
chrome_path = r'chromedriverの登載パス'
options = Options()
options.add_argument("--incognito")
driver = webdriver.Chrome(executable_path=chrome_path,options=options)

#調査URLへのアクセス
url = "調査URL"
driver.get(url)
sleep(1)

start=time()
print("スクロール開始")

#スクロール処理（スクロールしないと画面描画されないページの場合に実施）
height = driver.execute_script('return document.body.scrollHeight')
driver.execute_script("window.scrollTo(0,{});".format(height))
sleep(1)

#画面描画完了まで「ページの一番下へ遷移⇒画面描画⇒ページの一番下へ遷移・・・」を繰り返し
while True:
    driver.execute_script("window.scrollTo(0,{});".format(height))
    print("height:" + str(height))
    sleep(5)
    max_height = driver.execute_script('return document.body.scrollHeight')
    print("max_height:" + str(max_height))

    if height == max_height:
        break
    height = max_height

print("スクロール終了")
print("スクロールの処理時間：" + str(math.floor(time()-start)) + "秒です")

start2=time()

#BeautifulSoupにHTML情報を渡して解析開始
soup = BeautifulSoup(driver.page_source, 'lxml')
companys = soup.find_all('div', class_='bl_card2')
base_url = "調査URL(どのページでも共通の部分)"
d_list = []

for company in companys:
    print('='*50)
    company_names = company.select('span.bl_card2_ttl_text > a')
    if company_names == []:
        company_name = "" 
    else:
        company_name = company_names[0].text
    print(company_name)
    company_tels = company.select('div.telNo > p > strong > a')
    if company_tels == []:
        company_tel = "" 
    else:
        company_tel = company_tels[0].get('href')
    print(company_tel)
    page_urls = company.select('span.bl_card2_ttl_text > a')
    if page_urls == []:
        company_address = "" 
    else:
        page_url = base_url + page_urls[0].get('href')
        page_r = requests.get(page_url, timeout=3)
        page_r.raise_for_status()
 
        page_soup = BeautifulSoup(page_r.content, 'lxml')

        address = page_soup.find(text=re.compile('住所')).parent.parent
        company_address = address.find('p').text
        print(company_address)

    d = {
        "company_name": company_name,
        "company_tel": company_tel,
        "company_address": company_address
    }

    #取得結果をリストに格納
    d_list.append(d)
    sleep(1)

print('='*50)
print("全部で" + str(len(d_list)) + "件を取得しました。")
print("データ取得の処理時間：" + str(math.floor(time()-start2)) + "秒です")

#取得結果をcsvファイルに出力
df = pd.DataFrame(d_list)
df.to_csv('出力ファイル名.csv', index=False, encoding='utf-8-sig')
