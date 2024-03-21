from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import openpyxl


df = pd.read_excel("ファイルパスを入れてください")
INCI = df["INCI"].to_list()

df_brand = pd.read_excel('ファイルパスを入れてください', sheet_name='Brands')
Brand = df_brand["Brand List"].to_list()

INTERVAL = 2.5
INTERVAL_10 = 10

#options = Options()
#options.add_argument("--headless")
browser = webdriver.Chrome()
time.sleep(INTERVAL)

#HP上の全てのingredientsを取得し、INCIリストとの共通項のみ抽出
ALL_ingredients = []
browser.get("https://incidecoder.com/ingredients/all")

while True:
    if len(browser.find_elements(By.PARTIAL_LINK_TEXT, "Next page")) > 0:
        Find_name = browser.find_elements(By.CSS_SELECTOR, "div[class=paddingtbl] > a")
        for g in Find_name:
            ALL_ingredients.append(g.text)

        NEXT = browser.find_element(By.PARTIAL_LINK_TEXT, "Next page")
        NEXT.click()
        time.sleep(INTERVAL)

    else:
        break

matched_list = []
for inci in INCI:
    for all in ALL_ingredients:
        if inci == all:
            matched_list.append(inci)


#Excel
wb = openpyxl.Workbook()
sheet = wb.active
cnt = 1

#ingredientsから商品名とリンク取得のループ
browser.get("https://incidecoder.com")
time.sleep(INTERVAL)
action = webdriver.ActionChains(browser)

for B in Brand:
    browser.find_element(By.ID, "query").send_keys(B)
    browser.find_element(By.XPATH, "//div[1]/div[2]/div[1]/form/input[2]").click()
    time.sleep(INTERVAL)

    for I in matched_list:
        FIND_BOX = browser.find_element(By.XPATH, "//*[@id='products']/form/div[1]/div/div/span/span[1]/span/ul/li/input")
        PUT_ING = FIND_BOX.send_keys(I)
        time.sleep(INTERVAL)
        action.send_keys(Keys.ENTER).perform()

        FILTER = browser.find_element(By.XPATH, "//*[@id='products']/form/input[2]")
        FILTER.click()
        time.sleep(INTERVAL)

        browser.execute_script("window.scrollTo(0,document.body.scrollHeight);")
        time.sleep(INTERVAL)

#リコメンドがでても無視
        page_text = browser.find_element(By.TAG_NAME, "body").text
        search_text_1 = " instead."
        time.sleep(INTERVAL_10)

        if search_text_1 in page_text:
            time.sleep(INTERVAL_10)
            browser.execute_script("window.scrollTo(0,0);")
            time.sleep(INTERVAL)
            BACK_to_simple = browser.find_element(By.XPATH, "//*[@id='content']/div[2]/div/div[1]/a")
            BACK_to_simple.click()

        else:
            FIND_PLACE = browser.find_elements(By.CSS_SELECTOR, "div[class=std-side-padding] > a")
            for i in FIND_PLACE:
                print(i.text)
                print(i.get_attribute('href'))

                search_text_2 = ":("
                if search_text_2 in page_text:
                    break

                else:
                    cell1 = "A" + str(cnt)
                    sheet[cell1].value = B
                    cell2 = "B" + str(cnt)
                    sheet[cell2].value = I
                    cell3 = "C" + str(cnt)
                    sheet[cell3].value = i.get_attribute('href')
                    cnt += 1

            browser.execute_script("window.scrollTo(0,document.body.scrollHeight);")
            time.sleep(INTERVAL)

#Next Pageの処理
            while True:
                if len(browser.find_elements(By.PARTIAL_LINK_TEXT, "Next page")) >0:
                    NEXT = browser.find_element(By.PARTIAL_LINK_TEXT, "Next page")
                    NEXT.click()
                    time.sleep(INTERVAL)

                    FIND_PLACE = browser.find_elements(By.CSS_SELECTOR, "div[class=std-side-padding] > a")
                    for i in FIND_PLACE:
                        print(i.text)
                        print(i.get_attribute('href'))
                        time.sleep(INTERVAL)

                elif search_text_1 in page_text:
                    browser.execute_script("window.scrollTo(0,0);")
                    time.sleep(INTERVAL)
                    BACK_to_simple = browser.find_element(By.XPATH, "//*[@id='content']/div[2]/div/div[1]/a")
                    BACK_to_simple.click()


                else:
                    browser.execute_script("window.scrollTo(0,0);")
                    time.sleep(INTERVAL)
                    BACK_to_simple = browser.find_element(By.XPATH, "//*[@id='content']/div[2]/div/div[1]/a")
                    BACK_to_simple.click()
                    break

    Back_Home = browser.find_element(By.XPATH, "//*[@id='logo']")
    Back_Home.click()


wb.save("product_url.xlsx")
wb.close()