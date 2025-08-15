from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver.common.alert import Alert
from bs4 import BeautifulSoup
from datetime import datetime
import os
import time
import traceback
import csv
import pandas as pd
import sys

# ✅ 引数受け取り
if len(sys.argv) < 3:
    print("Usage: python announcement_info.py <Excelファイルパス> <保存先フォルダ> [開始日] [終了日]")
    sys.exit(1)

excel_path   = sys.argv[1]
download_dir = sys.argv[2]
start_date   = sys.argv[3] if len(sys.argv) > 3 else None
end_date     = sys.argv[4] if len(sys.argv) > 4 else None

print(f"📊 Excelファイル: {excel_path}")
print(f"📁 保存先: {download_dir}")
print(f"📅 渡された期間（未使用）: {start_date} ～ {end_date}")

# ✅ ログファイル初期化
log_path = os.path.join(download_dir, "found_seko_nos.csv")
if not os.path.exists(log_path):
    with open(log_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["施工番号", "種別", "ファイル名", "取得時刻"])

write_header = not os.path.exists("download_log.csv")
valid_exts = ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.zip']

# ✅ ダウンロード完了を待つ関数
def wait_for_download_complete(download_dir, filename, timeout=30):
    path = os.path.join(download_dir, filename)
    for _ in range(timeout):
        if os.path.exists(path) and not filename.endswith(".crdownload"):
            return True
        time.sleep(1)
    return False

# ✅ メイン処理
def download_bid_files(seko_nos, download_dir):
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_setting_values.popups": 0
    })
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument('--headless')

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

        found_seko_nos = set()

        while True:
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

            html = driver.page_source
            soup = BeautifulSoup(html, "html.parser")
            table = soup.find("table", {"id": "data"})
            rows = table.find("tbody").find_all("tr") if table else []

            for i, row in enumerate(rows):
                cells = row.find_all("td")
                if len(cells) >= 7:
                    seko_no = next(cells[0].stripped_strings)
                    if seko_no in seko_nos and seko_no not in found_seko_nos:
                        found_seko_nos.add(seko_no)
                        print(f"✅ 該当施工番号が見つかりました: {seko_no} → jsInfo({i}) 実行")

                        project_dir = os.path.join(download_dir, seko_no)
                        os.makedirs(project_dir, exist_ok=True)

                        driver.switch_to.default_content()
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))
                        driver.execute_script(f"jsInfo({i});")
                        time.sleep(2)

                        try:
                            WebDriverWait(driver, 3).until(EC.alert_is_present())
                            alert = Alert(driver)
                            print(f"⚠️ アラート検出: {alert.text}")
                            alert.accept()
                        except:
                            pass

                        driver.switch_to.default_content()
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
                        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmBOTTOM")))

                        detail_html = driver.page_source
                        with open(os.path.join(project_dir, f"detail_{seko_no}.html"), "w", encoding="utf-8") as f:
                            f.write(detail_html)

                        file_ids = ['01', '02', '03', '04', '05', '06', '10']
                        for fid in file_ids:
                            print(f"📥 ダウンロード開始: {fid}")
                            soup = BeautifulSoup(driver.page_source, "html.parser")
                            onclick_pattern = f"jsDoubleClick_Check({i}, '{fid}')"
                            found_link = soup.find("img", onclick=lambda x: x and onclick_pattern in x)

                            if not found_link:
                                print(f"🚫 ファイル {fid} のリンクが見つかりません → スキップ")
                                continue

                            before_files = set(os.listdir(download_dir))
                            driver.execute_script(f"jsDoubleClick_Check({i}, '{fid}');")
                            time.sleep(3)

                            downloaded_files = set(os.listdir(download_dir)) - before_files
                            for fname in downloaded_files:
                                if not fname.lower().endswith(tuple(valid_exts)):
                                    continue

                                if not wait_for_download_complete(download_dir, fname):
                                    print(f"⚠️ ダウンロード未完了: {fname}")
                                    continue

                                base, ext = os.path.splitext(fname)
                                new_name = f"{seko_no}_{fid}_{base}{ext}"
                                dst_path = os.path.join(project_dir, new_name)

                                if os.path.exists(dst_path):
                                    base, ext = os.path.splitext(new_name)
                                    new_name = f"{base}_{int(time.time())}{ext}"
                                    dst_path = os.path.join(project_dir, new_name)

                                try:
                                    os.rename(os.path.join(download_dir, fname), dst_path)
                                    print(f"✅ ファイル移動成功: {fname} → {new_name}")
                                    with open(log_path, "a", newline="", encoding="utf-8") as f:
                                        writer = csv.writer(f)
                                        writer.writerow([seko_no, fid, new_name, time.strftime("%Y-%m-%d %H:%M:%S")])
                                except Exception as e:
                                    print(f"❌ ファイル移動失敗: {fname} → {new_name} ({type(e).__name__})")

                            # クリーンアップ（施工番号関連ファイルのみ）
                            for fname in os.listdir(download_dir):
                                if fname.lower().endswith(tuple(valid_exts)) and seko_no in fname:
                                    try:
                                        os.remove(os.path.join(download_dir, fname))
                                    except Exception as e:
                                        print(f"⚠️ 削除失敗: {fname} ({type(e).__name__})")

            # ⏭️ 次ページへ
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

            with open(os.path.join(download_dir, "frmTOP_snapshot.html"), "w", encoding="utf-8") as f:
                f.write(driver.page_source)

            try:
                next_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "btnNextPage")))
                disabled = next_btn.get_attribute("disabled")
                if disabled is not None:
                    print("📄 最終ページに到達しました")
                    break
                print("⏭️ 次ページへ移動します")
                driver.execute_script("jsNextPage();")
                WebDriverWait(driver, 10).until(
                    lambda d: "次頁" not in d.page_source or "前頁" in d.page_source
                )
                time.sleep(2)
            except Exception as e:
                print(f"⚠️ 次ページボタンが見つかりません: {type(e).__name__}")
                break

        # ⚠️ 未取得施工番号の表示と保存
        not_found = set(seko_nos) - found_seko_nos
        if not_found:
            print("⚠️ 以下の施工番号は検索結果に見つかりませんでした:")
            with open(os.path.join(download_dir, "not_found_seko_nos.csv"), "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["未取得施工番号", "確認時刻"])
                for nf in not_found:
                    print(f" - {nf}")
                    writer.writerow([nf, time.strftime("%Y-%m-%d %H:%M:%S")])

    finally:
        driver.quit()
        print("🛑 ブラウザを終了しました")

# ✅ 実行ブロック
if __name__ == "__main__":
    df = pd.read_excel(excel_path)
    seko_nos = df["施工番号"].dropna().astype(str).tolist()

    print(f"📊 施工番号一覧: {seko_nos}")
    download_bid_files(seko_nos, download_dir)
    print("✅ 入札情報のダウンロードが完了しました。")