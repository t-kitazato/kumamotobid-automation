import logging
import re
import time
import pandas as pd
import sys
import os
import merge_bid_data
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options

# ✅ 引数受け取り
if len(sys.argv) < 3:
    print("Usage: python kumamotopre.py <保存先フォルダ> <Excelファイル> [開始日] [終了日]")
    sys.exit(1)

excel_file  = sys.argv[1]
save_folder = sys.argv[2]
start_date = sys.argv[3] if len(sys.argv) > 3 else None
end_date = sys.argv[4] if len(sys.argv) > 4 else None

print(f"📁 保存先: {save_folder}")
print(f"📋 Excelファイル: {excel_file}")
print(f"📅 渡された期間（未使用）: {start_date} ～ {end_date}")


# ✅ ログ設定
log_path = os.path.join(save_folder, "kumamoto_scraper.log")
logging.basicConfig(filename=log_path, level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")
logging.info("📌 スクリプト開始")

# ✅ 保存先フォルダの作成
os.makedirs(save_folder, exist_ok=True)

# ✅ ユーティリティ関数
def extract_text(soup, label):
    cell = soup.find("td", string=label)
    return cell.find_next("td").text.strip() if cell else "取得失敗"

def clean_price(text):
    match = re.search(r"\d[\d,]*円", text)
    return match.group(0) if match else text

def extract_location(soup):
    for label in ["工事場所", "場所", "施工場所"]:
        value = extract_text(soup, label)
        if value != "取得失敗":
            return value
    return "入札前"

def extract_award_info(soup):
    for row in soup.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) >= 3 and "落札" in cells[-1].text:
            return clean_price(cells[1].text.strip()), cells[0].text.strip()
    return "入札前", "入札前"

def contains_target_vendor(soup, target="北里道路（株）"):
    return target in soup.get_text()

# ✅ ブラウザ起動

options = Options()
headless_flag = os.environ.get("HEADLESS", "True")
if headless_flag == "True":
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(options=options)
driver.get("https://ebid.kumamoto-idc.pref.kumamoto.jp/PPIAccepter/TopServlet")
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//img[contains(@src, "logo_kumamoto_pref.gif")]'))).click()
time.sleep(1)

# 左メニューから検索画面へ
driver.switch_to.default_content()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmLEFT")))
driver.execute_script("jsLink(1,1);")
time.sleep(1)

# 検索条件指定
driver.switch_to.default_content()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmTOP")))

Select(driver.find_element(By.NAME, "NYUSATU_TYPE")).select_by_visible_text("通常型指名競争入札")
Select(driver.find_element(By.NAME, "GYOSYU_TYPE")).select_by_visible_text("工事")
Select(driver.find_element(By.NAME, "HACHU_TANTOU_KYOKU")).select_by_visible_text("阿蘇地域振興局")

driver.find_element(By.NAME, "btnSearch").click()
time.sleep(1)

# 📦 結果格納
bid_data = []
page_count = 0

# 🔁 ページ巡回
while True:
    page_count += 1
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

    rows = driver.find_elements(By.XPATH, '//table[@id="data"]/tbody/tr')
    logging.info(f"📄 ページ {page_count}: {len(rows)} 件の案件を検出")

    for index in range(len(rows)):
        try:
            rows = driver.find_elements(By.XPATH, '//table[@id="data"]/tbody/tr')
            img = rows[index].find_element(By.XPATH, './/img[contains(@onclick, "jsBidInfo")]')
            onclick_code = img.get_attribute("onclick").replace("javascript:", "").replace(";", "")
            driver.execute_script(onclick_code)
            logging.info(f"🖱️ jsBidInfo 実行: {onclick_code}")
            time.sleep(2)

            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmBOTTOM")))

            soup = BeautifulSoup(driver.page_source, "html.parser")
            if contains_target_vendor(soup):
                logging.info(f"✅ 指名業者に『北里道路（株）』が含まれています")
                bid_data.append([
                    extract_text(soup, "施行番号"),
                    extract_text(soup, "工事・業務名"),
                    "通常型指名競争入札",
                    extract_text(soup, "開札（予定）日"),
                    "",
                    clean_price(extract_text(soup, "予定価格")),
                    extract_location(soup),
                    *extract_award_info(soup)
                ])
            else:
                logging.info(f"⛔ 指名業者に『北里道路（株）』は含まれていません。スキップ")

            driver.execute_script("jsBack();")
            time.sleep(1)
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

        except Exception as e:
            logging.error(f"❌ 詳細取得エラー: {e}")

    # ✅ ページネーション処理
    try:
        driver.switch_to.default_content()
        WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
        WebDriverWait(driver, 5).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

        soup = BeautifulSoup(driver.page_source, "html.parser")
        next_button = soup.find("input", {"name": "btnNextPage", "type": "button", "value": "次頁"})
        if next_button and "disabled" not in next_button.attrs:
            logging.info("➡️ 次ページへ移動（jsNextPage 実行）")
            driver.execute_script("jsNextPage();")
            time.sleep(2)
        else:
            logging.info("📘 最終ページに到達。処理終了")
            break
    except Exception as e:
        logging.error(f"❌ ページネーションエラー: {e}")
        break

# 📤 Excel出力
output_path = os.path.join(save_folder, "北里道路_入札候補案件.xlsx")
df = pd.DataFrame(bid_data, columns=[
    "施工番号", "工事・業務名", "入札及び契約方法", "開札予定日",
        "備考", "予定価格", "工事場所", "落札金額", "落札業者"
])
df.to_excel(output_path, index=False)
logging.info(f"📦 Excel保存完了: {output_path}")

# 🛑 終了処理
driver.quit()
logging.info("🛑 ブラウザを終了しました")
print("✅ 処理が完了しました。")

#　入札管理.xlsx へ反映
existing_excel = "入札管理data.xlsx"
new_excel = output_path  # 取得結果のExcelパス
merge_bid_data.merge_preserving_format(existing_excel, new_excel)

