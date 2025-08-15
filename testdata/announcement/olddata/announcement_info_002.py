from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.common.alert import Alert
import os
import time
import traceback

# ✅ ダウンロード先の設定（パスは raw 文字列で）
download_dir = r"D:\work\kumamotopre\announcement"

# ✅ ダウンロード完了を待つ関数
def wait_for_downloads(folder, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        if not any(fname.endswith(".crdownload") for fname in os.listdir(folder)):
            return True
        time.sleep(1)
    return False

# ✅ Chromeオプション設定
options = Options()
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1  # ✅ これが重要
})

driver = webdriver.Chrome(options=options)

try:
    # 🌐 トップページへ
    driver.get("https://ebid.kumamoto-idc.pref.kumamoto.jp/PPIAccepter/TopServlet")
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//img[contains(@src, "logo_kumamoto_pref.gif")]'))
    ).click()
    time.sleep(3)

    # 📋 左メニューから検索画面へ
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmLEFT")))
    driver.execute_script("jsLink(3,1);")
    time.sleep(3)

    # 🔍 検索条件指定
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmTOP")))

    Select(driver.find_element(By.NAME, "GYOSYU_TYPE")).select_by_visible_text("工事")
    Select(driver.find_element(By.NAME, "HACHU_TANTOU_KYOKU")).select_by_visible_text("阿蘇地域振興局")

    driver.find_element(By.NAME, "btnSearch").click()
    time.sleep(2)

    # 📄 検索結果画面へ
    driver.switch_to.default_content()
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table", {"id": "data"})
    rows = table.find("tbody").find_all("tr") if table else []

    print("✅ 検索結果画面HTML取得完了")

    # 🎯 施工番号でフィルタして詳細ページへ
    target_seko_nos = ["2525-5070260213", "2525-5070260198", "2525-5070260130"]
    found_seko_nos = set()
    
    for i, row in enumerate(rows):
        cells = row.find_all("td")
        if len(cells) >= 7:
            seko_no = next(cells[0].stripped_strings)
            if seko_no in target_seko_nos and seko_no not in found_seko_nos:
                found_seko_nos.add(seko_no)
                print(f"✅ 該当施工番号が見つかりました: {seko_no} → jsInfo({i}) 実行")
                driver.execute_script("return typeof jsInfo === 'function';")

                try:
                    driver.execute_script(f"jsInfo({i});")
                    time.sleep(2)

                    # アラート処理
                    try:
                        WebDriverWait(driver, 3).until(EC.alert_is_present())
                        alert = Alert(driver)
                        print(f"⚠️ アラート検出: {alert.text}")
                        alert.accept()
                        continue  # 次の行へスキップ
                    except:
                            pass

                    # 📄 詳細処理（そのままでOK）
                    
                    driver.switch_to.default_content()
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmBOTTOM")))
                                      
                    # 📥 ファイル種別ごとにダウンロード
                    file_ids = ['01', '02', '03', '04', '05', '06', '10']
                    for fid in file_ids:
                        print(f"📥 ダウンロード開始: {fid}")
                        driver.execute_script(f"jsDoubleClick_Check(0, '{fid}');")
                        
                        if wait_for_downloads(download_dir, timeout=30):
                            print(f"✅ ダウンロード完了: {fid}")
                        else:
                            print(f"⚠️ ダウンロードタイムアウト: {fid}")
                            
                    project_dir = os.path.join(download_dir, seko_no)
                    os.makedirs(project_dir, exist_ok=True)

                    for fname in os.listdir(download_dir):
                        if fname.endswith(".pdf"):  # 例
                            os.rename(
                                os.path.join(download_dir, fname),
                                os.path.join(project_dir, fname)
                            )

                    # 💾 詳細HTML保存
                    detail_html = driver.page_source
                    with open(os.path.join(download_dir, f"detail_{seko_no}.html"), "w", encoding="utf-8") as f:
                        f.write(detail_html)
                                                    
                    downloaded_files = os.listdir(download_dir)
                    print("📥 ダウンロードされたファイル一覧:")
                    for f in downloaded_files:
                        print(f" - {f}")
                  
                    print(f"📄 詳細ページ保存完了: detail_{seko_no}.html")
                    continue

                except UnexpectedAlertPresentException as e:
                    print(f"❌ アラート例外: {e}")
                    try:
                        alert = Alert(driver)
                        alert.accept()
                    except:
                        pass
                    continue
                        
    # 見つからなかった施工番号の表示
    not_found = set(target_seko_nos) - found_seko_nos
    if not_found:
        print("⚠️ 以下の施工番号は検索結果に見つかりませんでした:")
        for nf in not_found:
            print(f"⚠️ 該当施工番号が検索結果に見つかりませんでした: {target_seko_nos}")
            for nf in not_found:
                print(f" - {nf}")
                            
except Exception as e:
    print(f"❌ エラー発生: {type(e).__name__}: {e}")
    traceback.print_exc()
    with open("error_page.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)

finally:
    driver.quit()
