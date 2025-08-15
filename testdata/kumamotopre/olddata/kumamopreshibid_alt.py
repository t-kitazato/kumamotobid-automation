import logging
import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException
from bs4 import BeautifulSoup

# ✅ ログ設定
logging.basicConfig(
    filename="kumamoto_scraper.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("📌 スクリプト開始")

# 🧠 ユーティリティ関数
def extract_text(soup, label):
    cell = soup.find("td", string=label)
    return cell.find_next("td").text.strip() if cell else "取得失敗"

def clean_price(text):
    match = re.search(r"\d[\d,]*円", text)
    return match.group(0) if match else text

def extract_location(soup):
    keywords = ["工事場所", "場所", "施工場所"]
    for label in keywords:
        value = extract_text(soup, label)
        if value != "取得失敗":
            return value
    return "取得失敗"

def extract_award_info(soup):
    rows = soup.find_all("tr")
    for row in rows:
        cells = row.find_all("td")
        if len(cells) >= 3 and "落札" in cells[-1].text:
            金額 = clean_price(cells[1].text.strip())
            業者 = cells[0].text.strip()
            return 金額, 業者
    return "取得失敗", "取得失敗"

# 🌐 ブラウザ起動
driver = webdriver.Chrome()

try:
    driver.get("https://ebid.kumamoto-idc.pref.kumamoto.jp/PPIAccepter/TopServlet")
    logging.info("✅ トップページにアクセスしました")
    time.sleep(2)

    logo = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//img[contains(@src, "logo_kumamoto_pref.gif")]'))
    )
    logo.click()
    logging.info("✅ 熊本県ロゴをクリックしました")
    time.sleep(1)

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmLEFT")))
    driver.execute_script("jsLink(1,1);")
    logging.info("✅ frmLEFT 内で jsLink(1,1) を実行しました")
    time.sleep(1)

    # frmTOP フレーム内で検索条件を指定
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmTOP")))

    # 検索ページのHTMLを保存
    with open(f"debug_search.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    logging.info("📄 検索ページHTMLを保存しました: debug_search_0.html")


    # 入札及び契約の方法
    method_select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "NYUSATU_TYPE"))
    )
    method_select.send_keys("通常型指名競争入札")

    # 業種分類
    industry_select = driver.find_element(By.NAME, "GYOSYU_TYPE")
    industry_select.send_keys("工事")

    # 発注担当部局等
    dept_select = driver.find_element(By.NAME, "HACHU_TANTOU_KYOKU")
    dept_select.send_keys("阿蘇地域振興局")


    # 検索ボタンをクリック
    search_button = driver.find_element(By.NAME, "btnSearch")
    driver.execute_script("arguments[0].click();", search_button)

    # 検索結果テーブルの読み込みを待機
    #WebDriverWait(driver, 10).until(
    #    EC.presence_of_element_located((By.XPATH, '//table[@id="data"]/tbody/tr'))
    #)
    #rows = driver.find_elements(By.XPATH, '//table[@id="data"]/tbody/tr')

    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//input[@name="btnSearch"]'))
    )
    driver.execute_script("arguments[0].click();", search_button)
    logging.info("✅ 検索ボタンをクリックしました")
    time.sleep(1)

    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

    rows = driver.find_elements(By.XPATH, '//table[@id="data"]/tbody/tr')
    logging.info(f"🔍 検索結果から {len(rows)} 件の行を検出")
    
    with open("debug_search_results.html", "w", encoding="utf-8") as f:
         f.write(driver.page_source)
    logging.info("📝 検索結果HTMLを保存しました: debug_search_results.html")

    案件リスト = []
    for index, row in enumerate(rows):
        try:
            img = row.find_element(By.XPATH, './/img[contains(@onclick, "jsBidInfo")]')
            onclick_code = img.get_attribute("onclick")
            cells = row.find_elements(By.TAG_NAME, "td")
            logging.info(f"🔍 [{index}] セル数: {len(cells)}")

            if len(cells) >= 10:
                案件 = {
                    "onclick": onclick_code,
                    "施行番号": cells[0].text.strip().split("\n")[0],  # 上段の.
                    "工事名": cells[2].text.strip(),
                    "契約方法": cells[3].text.strip(),
                    "開札予定日": cells[4].text.strip(),
                    "備考": cells[9].text.strip()
                }
                案件リスト.append(案件)
                logging.info(f"📥 [{index}] 案件をリストに追加: {案件['施行番号']}")
                logging.info(f"🔗 onclickコード: {案件['onclick']}")
                logging.info(f"📥 [{index}] 案件追加: {案件['施行番号']}")

            else:
                logging.warning(f"⚠️ [{index}] セル数不足でスキップ")
        except Exception as e:
            logging.error(f"❌ [{index}] 案件取得エラー: {e}")
    logging.info(f"📊 案件リスト完成: {len(案件リスト)} 件")
    
    bid_data = []
    for index, 案件 in enumerate(案件リスト):
        try:
            onclick_code = 案件["onclick"].strip().replace("javascript:", "").replace(";", "")

            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
            logging.info("🧭 frmRIGHT に切り替えました")
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))
            logging.info("🧭 frmMIDDLE に切り替えました")
            logging.info("🔄 検索結果ページに復帰しました")

            # BidInfoアイコンをクリック（containsで柔軟に）
            detail_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f'//img[@class="ATYPE" and contains(@onclick, "{onclick_code}")]'))
            )
            detail_button.click()
            logging.info(f"🖱️ jsBidInfo 実行: {onclick_code}")

            # 少し待機（JavaScriptでフレームが切り替わるまで）
            time.sleep(2)
            
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmBOTTOM")))

            # 詳細ページのHTMLを保存
            with open(f"debug_detail_{index}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            logging.info("📄 詳細ページHTMLを保存しました: debug_detail_0.html")
            
            soup = BeautifulSoup(driver.page_source, "html.parser")

            raw_price = extract_text(soup, "予定価格")
            予定価格 = clean_price(raw_price)
            工事場所 = extract_location(soup)
            落札金額, 落札業者 = extract_award_info(soup)
            
            if 落札金額 == "取得失敗":
                落札金額 = "落札なし"
                落札業者 = "該当なし"

            bid_data.append([
                案件["施行番号"], 案件["工事名"], 案件["契約方法"], 案件["開札予定日"],
                案件["備考"], 予定価格, 工事場所, 落札金額, 落札業者
            ])
            logging.info(f"✅ [{index}] 詳細情報取得完了: {案件['施行番号']}")
            driver.execute_script("jsBack();")
            logging.info("🔙 jsBack() を実行して検索結果一覧に戻りました")
            time.sleep(1)

        except Exception as e:
            logging.error(f"❌ [{index}] 詳細情報取得エラー: {e}")

    df = pd.DataFrame(bid_data, columns=[
        "施工番号", "工事・業務名", "入札及び契約方法", "開札予定日",
        "備考", "予定価格", "工事場所", "落札金額", "落札業者"
    ])
    df.to_excel("kumamoto_bids_detailed_revised.xlsx", index=False)
    logging.info("📦 Excel保存完了: kumamoto_bids_detailed_revised.xlsx")

except TimeoutException as e:
    logging.error(f"⏱️ 要素取得に失敗しました: {e}")

except UnexpectedAlertPresentException:
    alert = Alert(driver)
    logging.warning(f"⚠️ アラート検知: {alert.text}")
    alert.accept()
    logging.info("✅ アラートを閉じました")

finally:
    driver.quit()
    logging.info("🛑 ブラウザを終了しました")
