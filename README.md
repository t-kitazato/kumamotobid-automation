📘 README: 入札情報取得ツール（GUI版）
🛠 概要
このツールは、熊本県をはじめとする自治体の入札情報を自動収集・処理するPythonベースのアプリケーションです。tkinterによるGUIを備え、非エンジニアでも簡単に操作可能。Seleniumによるブラウザ操作、Excel連携、日付指定、スクリプト選択など、柔軟な運用が可能です。

🎯 主な機能
機能	説明
📁 保存先フォルダ選択	ダウンロードファイルの保存先を指定
📋 Excelファイル選択	施工番号などの情報を含むExcelを指定
📅 日付指定	開始日・終了日をカレンダーから選択
🐍 スクリプト選択	kumamotopre.py（入札情報更新）または announcement_info.py（仕様書取得）を選択
✅ ヘッドレスモード切替	ブラウザ表示のON/OFFをチェックボックスで切替
▶️ 実行ボタン	指定条件でPythonスクリプトを実行
📦 必要な環境
Python 3.9+

以下のライブラリ（pip installでインストール可能）:

bash
pip install selenium pandas openpyxl tkcalendar
ChromeDriver（Selenium用）を環境に合わせて配置

📂 ファイル構成例
Code
project/
├─ kumamotobid_main.py         # GUI本体
├─ kumamotopre.py          # 入札情報取得スクリプト
├─ announcement_info.py    # 仕様書ダウンロードスクリプト
├─ 北里道路_入札候補案件.xlsx  # Excel入力例
├─ logs/                   # 実行ログ保存先（任意）
└─ downloads/              # ダウンロードファイル保存先
🚀 実行方法
bash
python kumamotobid_main.py
GUIが起動し、各種パラメータを設定後「▶️ 実行」ボタンを押すことで、指定スクリプトが実行されます。

🔧 スクリプト引数仕様（内部処理）
各スクリプトは以下の引数を受け取ります：

bash
python kumamotopre.py <Excelファイル> <開始日> <終了日> <保存フォルダ>
環境変数 HEADLESS によりブラウザ表示の有無を制御可能：

bash
HEADLESS=True python kumamotopre.py ...

📈 今後の展望
実行ログのGUI表示（Textウィジェット）

進捗バーの追加

スクリプトの選択肢拡張（市町村別など）

GitHub連携によるバージョン管理と更新通知

マルチPC展開用のバッチインストーラー

👤 作者
t（熊本県南小国町）
