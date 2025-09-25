# 🕒 Working Hours Calculation Application

日々の勤務時間（出勤・退勤・休憩）を記録し、**当日／週間／月間**の労働時間や**概算賃金**を集計できる Windows 向け **GUI アプリ**です。  
実行ファイル（.exe）で配布可能・インストール不要のポータブル運用にも対応します。

---

## ✨ 特長

- 直感的な **GUI 操作**（出勤／退勤／休憩を入力 → 集計一発）
- **当日／週／月**の勤務時間を自動集計
- **時給計算**（時間 × 単価）の簡易サポート
- **データ保存／読み込み**（JSON）
- **単体 .exe** 配布に対応（PyInstaller）
- Visual Studio（`.sln` / `.pyproj`）での開発にも対応

---

## 🖼 画面構成（Canvas）詳細

> ※ スクリーンショットの置き場所の例：`docs/screenshot_main.png`  
> 画像が用意でき次第、以下の `![...]()` を有効化してください。

<!--
![メイン画面](docs/screenshot_main.png)
-->

**メイン（入力キャンバス）**
- **日付選択**：対象日を指定します。
- **出勤時刻 / 退勤時刻**：時刻入力欄。キーボード／上下矢印で調整可能。
- **休憩時間**：分単位または時刻区間で入力（実装仕様に合わせて運用）。
- **時給**：任意。入力すると当日・期間集計に賃金概算を表示。
- **保存 / 読み込み**：`work_data.json` への保存、既存データの再読込。
- **集計表示（右ペイン）**：当日・今週・今月の合計勤務時間／（任意）賃金概算。

**メニュー／操作（例）**
- **ファイル**：新規、開く、保存、終了
- **編集**：当日入力のクリア、前日／翌日へ移動
- **表示**：週／月の範囲切替、集計の再計算
- **ヘルプ**：バージョン情報

> UI の文言や項目名は実装に合わせて読み替えてください。README は「画面の意図」が伝わることを重視しています。

---

## 💾 データ保存（プライバシー）

- データは既定で **`Working hours calculation application/work_data.json`** に保存されます。  
- バックアップ **`work_data.bak`** を生成する場合があります。  
- これらは **個人データ** です。公開リポジトリでは **Git 管理対象外** にしてください（`.gitignore` で除外済み）。

> 既にコミットしてしまった場合は `git rm --cached` で追跡解除してください（履歴から完全に消す場合は `git filter-repo` や BFG の利用をご検討ください）。

---

## 📥 ユーザー向けインストール（.exe 版）

1. **Releases ページ** から最新の実行ファイルをダウンロード  
   `https://github.com/totsuka0405/Working_hours_calculation_application/releases`
2. 任意のフォルダに `Working_hours_calculation_application.exe` を配置
3. ダブルクリックで起動（インストール不要）

> SmartScreen などでブロックされた場合は「詳細情報」→「実行」を選択してください。配布元が個人署名であるため初回は警告が出ることがあります。

---

## 💻 開発者向け（ソースから実行）

```bash
git clone https://github.com/totsuka0405/Working_hours_calculation_application.git
cd Working_hours_calculation_application

# 必要に応じて仮想環境を作成
# python -m venv .venv && .venv\Scripts\activate

pip install -r requirements.txt
python Working_hours_calculation_application/Working_hours_calculation_application.py
