from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

download_dir = r"D:\work\kumamotopre\announcement"

options = Options()

options.add_argument("--headless")
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(options=options)


# Chromeのオプション設定（ヘッドレスモードなど）
# options = Options()
# options.add_argument("--headless")
# driver = webdriver.Chrome(options=options)

try:
    # 🌐 ブラウザ起動
    driver = webdriver.Chrome()
    driver.get("https://ebid.kumamoto-idc.pref.kumamoto.jp/PPIAccepter/TopServlet")
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//img[contains(@src, "logo_kumamoto_pref.gif")]'))).click()
    time.sleep(3)

    # 左メニューから検索画面へ
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmLEFT")))
    driver.execute_script("jsLink(3,1);")
    time.sleep(3)
    
    # 検索条件指定
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmTOP")))

    # HTML取得
    html = driver.page_source
    with open("announcement_search.html", "w", encoding="utf-8") as f:
        f.write(html)
        
    print("✅ 検索画面HTML取得完了")

    Select(driver.find_element(By.NAME, "GYOSYU_TYPE")).select_by_visible_text("工事")
    Select(driver.find_element(By.NAME, "HACHU_TANTOU_KYOKU")).select_by_visible_text("阿蘇地域振興局")

    driver.find_element(By.NAME, "btnSearch").click()
    time.sleep(1)

    # 検索結果画面へ移動
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

    # HTML取得
    html = driver.page_source
    with open("announcement_search_result.html", "w", encoding="utf-8") as f:
        f.write(html)
        # BeautifulSoupで解析
        soup = BeautifulSoup(html, "html.parser")

    table = soup.find("table", {"id": "data"})
    rows = table.find("tbody").find_all("tr") if table else []

    for i, row in enumerate(rows):
        cells = row.find_all("td")
        if len(cells) >= 7:
            seko_no = cells[0].get_text(strip=True).split("\n")[0]
            gyosyu = cells[1].get_text(strip=True)
            kouji_name = cells[3].get_text(strip=True)
            nyusatsu_type = cells[4].get_text(strip=True)
            kokoku_date = cells[6].get_text(strip=True)
            print(f"🔹 [{i}] {seko_no} | {gyosyu} | {kouji_name} | {nyusatsu_type} | {kokoku_date}")

    print("✅ 検索結果画面HTML取得完了")
    
    # 🔍 施工番号でフィルタして詳細ページへ遷移
    target_seko_no = "2525-5070260213"  # ← 任意の施工番号に変更可能

    for i, row in enumerate(rows):
        cells = row.find_all("td")
        if len(cells) >= 7:
            seko_no = next(cells[0].stripped_strings)
            if seko_no == target_seko_no:
                print(f"✅ 該当施工番号見つかりました: {seko_no} → jsInfo({i}) 実行")

                # JavaScriptで詳細ページ表示
                driver.execute_script(f"jsInfo({i});")
                time.sleep(2)

                # フレーム切り替え（frmBOTTOMに詳細が表示される）
                driver.switch_to.default_content()
                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmBOTTOM")))

                file_ids = ['01', '02', '03', '04', '05', '06', '10']
                for fid in file_ids:
                    driver.execute_script(f"jsDownload('{fid}')")
                    time.sleep(2)  # ダウンロード待機（必要に応じて調整）


                # 詳細HTML取得・保存
                detail_html = driver.page_source
                with open(f"detail_{seko_no}.html", "w", encoding="utf-8") as f:
                    f.write(detail_html)
                    print(detail_html[:1000])

                print(f"📄 詳細ページ保存完了: detail_{seko_no}.html")
                break
    else:
        print(f"⚠️ 該当施工番号が検索結果に見つかりませんでした: {target_seko_no}")


finally:
    driver.quit()
