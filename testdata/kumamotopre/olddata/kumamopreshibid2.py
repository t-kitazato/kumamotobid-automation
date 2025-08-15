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

def main():
    driver.get("https://ebid.kumamoto-idc.pref.kumamoto.jp/PPIAccepter/TopServlet")
    logging.info("✅ トップページにアクセスしました")
    time.sleep(2)

    # 熊本県ロゴクリック
    switch_to_frame(driver, ["frmLEFT"])
    driver.execute_script("jsLink(1,1)")
    logging.info("✅ 熊本県ロゴをクリックしました")

    # 検索ボタンクリック
    switch_to_frame(driver, ["frmRIGHT", "frmMIDDLE"])
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "btnSearch"))).click()
    logging.info("✅ 検索ボタンをクリックしました")

    # 検索結果取得
    rows = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//tr[contains(@class, "list")]'))
    )
    logging.info(f"🔍 検索結果から {len(rows)} 件の行を検出")

    案件リスト = []
    for index, row in enumerate(rows):
        try:
            img = row.find_element(By.XPATH, './/img[contains(@onclick, "jsBidInfo")]')
            onclick_code = img.get_attribute("onclick")
            案件リスト.append(onclick_code)
        except Exception as e:
            logging.warning(f"⚠️ [{index}] jsBidInfo取得失敗: {e}")
            continue

    logging.info(f"📋 案件リストの件数: {len(案件リスト)}")

    bid_data = []
    for i, onclick_code in enumerate(案件リスト):
        try:
            # フレーム切り替えして詳細ページへ
            switch_to_frame(driver, ["frmRIGHT", "frmMIDDLE"])
            driver.execute_script(onclick_code)
            logging.info(f"➡️ [{i}] jsBidInfo 実行: {onclick_code}")
            time.sleep(1.5)  # ページ遷移待ち

            # 詳細情報取得（例: 工事場所、落札金額、落札業者）
            switch_to_frame(driver, ["frmRIGHT", "frmBOTTOM"])
            工事場所 = driver.find_element(By.ID, "lblPlace").text
            落札金額 = driver.find_element(By.ID, "lblBidPrice").text
            落札業者 = driver.find_element(By.ID, "lblBidder").text

            bid_data.append({
                "工事場所": 工事場所,
                "落札金額": 落札金額,
                "落札業者": 落札業者
            })
            logging.info(f"✅ [{i}] 詳細取得成功: {工事場所}, {落札金額}, {落札業者}")

        except Exception as e:
            logging.warning(f"❌ [{i}] 詳細取得失敗: {e}")
            continue

    # Excel保存
    df = pd.DataFrame(bid_data)
    df.to_excel("kumamoto_bids_detailed_revised.xlsx", index=False)
    logging.info("📦 Excel保存完了: kumamoto_bids_detailed_revised.xlsx")

    driver.quit()
    logging.info("🛑 ブラウザを終了しました")

if __name__ == "__main__":
    main()
