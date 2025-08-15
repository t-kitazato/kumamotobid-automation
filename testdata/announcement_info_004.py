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
import csv
import requests
import sys
import pandas as pd
import uuid

# ✅ ダウンロードディレクトリ
download_dir = r"D:\work\kumamotobid\announcement"

valid_exts = ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.zip']

write_header = not os.path.exists("download_log.csv")

# ✅ ダウンロード完了を待つ関数
def wait_for_downloads(folder, timeout=30):
    start = time.time()
    while time.time() - start < timeout:
        if not any(fname.endswith(".crdownload") for fname in os.listdir(folder)):
            return True
        time.sleep(1)
    return False

def wait_for_download_complete(download_dir, filename, timeout=30):
    path = os.path.join(download_dir, filename)
    for _ in range(timeout):
        if os.path.exists(path) and not filename.endswith(".crdownload"):
            return True
        time.sleep(1)
    return False

def download_bid_files(seko_nos, download_dir):
   # ✅ Chromeオプション設定
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
        
        # 🎯 施工番号ごとに処理
        target_seko_nos = seko_nos
        found_seko_nos = set()

        for target_seko_no in target_seko_nos:
            while True:  # ページ送りループ
                # 📄 検索結果画面へ
                driver.switch_to.default_content()
                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmMIDDLE")))

                html = driver.page_source
                soup = BeautifulSoup(html, "html.parser")
                table = soup.find("table", {"id": "data"})
                rows = table.find("tbody").find_all("tr") if table else []

                # ページ内の施工番号一覧を取得
                for i, row in enumerate(rows):
                    cells = row.find_all("td")
                    if len(cells) >= 7:
                        seko_no = next(cells[0].stripped_strings)
                        if seko_no in seko_nos and seko_no not in found_seko_nos:
                            found_seko_nos.add(seko_no)
                            print(f"✅ 該当施工番号が見つかりました: {seko_no} → jsInfo({i}) 実行")

                            project_dir = os.path.join(download_dir, seko_no)
                            os.makedirs(project_dir, exist_ok=True)
                            print(f"📂 フォルダ作成: {project_dir}")

                            driver.execute_script(f"jsInfo({i});")
                            time.sleep(2)

                            # アラート処理
                            try:
                                WebDriverWait(driver, 3).until(EC.alert_is_present())
                                alert = Alert(driver)
                                print(f"⚠️ アラート検出: {alert.text}")
                                alert.accept()
                                break  # 次の施工番号へ
                            except:
                                pass

                            # 📄 詳細画面へ
                            driver.switch_to.default_content()
                            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
                            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmBOTTOM")))

                            # 💾 詳細HTML保存
                            detail_html = driver.page_source
                            with open(os.path.join(project_dir, f"detail_{seko_no}.html"), "w", encoding="utf-8") as f:
                                f.write(detail_html)
                            print(f"📄 詳細ページ保存完了: detail_{seko_no}.html")
                            print("📂 現在のダウンロードディレクトリの内容:")
                            downloaded_files = os.listdir(download_dir)
                            for fname in os.listdir(download_dir):
                                if not fname.lower().endswith(tuple(valid_exts)):
                                    continue
                                
                                src_path = os.path.join(download_dir, fname)
                                base, ext = os.path.splitext(fname)
                                new_name = f"{seko_no}_{fid}_{uuid.uuid4().hex[:8]}{ext}"
                                dst_path = os.path.join(project_dir, new_name)
                                
                                if not wait_for_download_complete(download_dir, fname):
                                    print(f"⚠️ ダウンロード未完了: {fname}")
                                    continue
                                
                                # ファイル名衝突を避ける
                                if os.path.exists(dst_path):
                                    base, ext = os.path.splitext(new_name)
                                    new_name = f"{base}_{int(time.time())}{ext}"
                                    dst_path = os.path.join(project_dir, new_name)
                                    print(f"📦 移動先: {dst_path}")
                                try:
                                    os.rename(src_path, dst_path)
                                    
                                    print(f"✅ ファイル移動成功: {fname} → {new_name}")
                                except Exception as e:
                                    print(f"❌ ファイル移動失敗: {fname} → {new_name} ({type(e).__name__})")
                                try:
                                    os.remove(os.path.join(download_dir, fname))
                                except Exception as e:
                                    print(f"⚠️ 削除失敗: {fname} ({type(e).__name__})")

                            # 📥 ファイル種別ごとにダウンロード
                            file_ids = ['01', '02', '03', '04', '05', '06', '10']                     
                            for fname in os.listdir(download_dir):
                                if not fname.lower().endswith(tuple(valid_exts)):
                                    continue
                            for fid in file_ids:
                                print(f"📥 ダウンロード開始: {fid}")

                                # frmBOTTOM の HTML を再取得してリンク確認
                                soup = BeautifulSoup(driver.page_source, "html.parser")
                                # found_link = soup.find("img", onclick=lambda x: x and f"jsDoubleClick_Check(0, '{fid}')" in x)
                                onclick_pattern = f"jsDoubleClick_Check({i}, '{fid}')"
                                found_link = soup.find("img", onclick=lambda x: x and onclick_pattern in x)

                                if not found_link:
                                    print(f"🚫 ファイル {fid} のリンクが見つかりません → スキップ")
                                    continue

                                # ダウンロード前のファイル一覧
                                before_files = set(os.listdir(download_dir))

                                # JavaScript 実行でダウンロード開始
                                driver.execute_script(f"jsDoubleClick_Check({i}, '{fid}');")
                                print(f"📥 ダウンロード対象: {fname}")

                                try:
                                    WebDriverWait(driver, 10).until(
                                        lambda d: any(
                                            fname.lower().endswith((".pdf", ".zip", ".xls", ".doc", ".xlsx"))
                                            for fname in os.listdir(download_dir)
                                        )
                                    )
                                    src = driver.find_element(By.NAME, "ifrmDownload").get_attribute("src")
                                    print(f"📎 iframe src for {fid}: {src}")
                                except:
                                    print(f"⚠️ iframe src 未取得: {fid}")

                                # ダウンロード完了確認（.crdownload に依存しない方法も併用）
                                time.sleep(3)
                                downloaded_files = os.listdir(download_dir)
                                for fname in downloaded_files:
                                    if fname.lower().endswith((".pdf", ".zip", ".xls", ".doc", ".xlsx")):
                                        base, ext = os.path.splitext(fname)
                                        # new_name = f"{seko_no}_{fid}_{uuid.uuid4().hex[:8]}{ext}"
                                        new_name = f"{seko_no}_{fid}_{base}{ext}"

                                        new_path = os.path.join(project_dir, new_name)

                                        # ✅ ダウンロード完了を待機
                                        if not wait_for_download_complete(download_dir, fname):
                                            print(f"⚠️ ダウンロード未完了（crdownload残り）: {fname}")
                                            continue

                                        # ✅ ファイル移動
                                        if os.path.exists(new_path):
                                            base, ext = os.path.splitext(new_name)
                                            new_name = f"{base}_{int(time.time())}{ext}"
                                            new_path = os.path.join(project_dir, new_name)

                                        try:
                                            os.rename(os.path.join(download_dir, fname), new_path)
                                            print(f"✅ ダウンロード完了: {fid} → {new_name}")
                                            
                                            global write_header
                                            with open("download_log.csv", "a", newline="", encoding="utf-8") as f:
                                                writer = csv.writer(f)
                                                if write_header:
                                                    writer.writerow(["施工番号", "種別", "ファイル名", "取得時刻"])
                                                writer.writerow([seko_no, fid, new_name, time.strftime("%Y-%m-%d %H:%M:%S")])
                                        except Exception as e:
                                            print(f"❌ ファイル移動失敗: {fname} → {new_name} ({type(e).__name__})")
                                            continue  # 次のファイルへ

                                for fname in os.listdir(download_dir):
                                    if fname.lower().endswith((".pdf", ".zip", ".xls", ".doc", ".xlsx")):
                                        os.remove(os.path.join(download_dir, fname))
                driver.switch_to.default_content()
                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmRIGHT")))
                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "frmTOP")))
                try:
                    next_btn = driver.find_element(By.NAME, "btnNextPage")
                    if next_btn.get_attribute("disabled"):
                        print("📄 最終ページに到達しました")
                        break
                    driver.execute_script("jsNextPage();")
                    time.sleep(2)
                except
                    print("⚠️ 次ページボタンが見つかりません")
                    break

                except Exception as e:
                    print(f"❌ エラー発生（{target_seko_no}）: {type(e).__name__}: {e}")
                    traceback.print_exc()
                    with open(os.path.join(download_dir, f"error_{target_seko_no}.html"), "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                    continue
                break
        # 見つからなかった施工番号の表示
        not_found = set(target_seko_nos) - found_seko_nos
        if not_found:
            print("⚠️ 以下の施工番号は検索結果に見つかりませんでした:")
            for nf in not_found:
                print(f" - {nf}")

    finally:
        driver.quit()
        
if __name__ == "__main__": 
    import pandas as pd
    excel_path = r"D:\work\kumamotobid\北里道路_入札候補案件.xlsx"
    df = pd.read_excel(excel_path)
    seko_nos = df["施工番号"].dropna().astype(str).tolist()
    
    print(f"📊 施工番号一覧: {seko_nos}")
    print(f"📁 ダウンロード先: {download_dir}")

    download_bid_files(seko_nos, download_dir)
