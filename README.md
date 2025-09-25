# 🕒 Working Hours Calculation Application

日々の勤務時間（出勤・退勤・休憩）を記録し、**当日／週間／月間**の労働時間や**概算賃金**を集計できる Windows 向け **GUI アプリ**です。
実行ファイル（.exe）で配布可能・インストール不要のポータブル運用にも対応します。

---

## ✨ 特長

* 直感的な **GUI 操作**（出勤／退勤／休憩を入力 → 集計一発）
* **当日／週／月**の勤務時間を自動集計
* **時給計算**（時間 × 単価）の簡易サポート
* **データ保存／読み込み**（JSON）
* **単体 .exe** 配布に対応（PyInstaller）
* Visual Studio（`.sln` / `.pyproj`）での開発にも対応

---

## 🖼 画面構成（Canvas）詳細

**メイン（入力キャンバス）**

* **日付選択**：対象日を指定します。
* **出勤時刻 / 退勤時刻**：時刻入力欄。キーボード／上下矢印で調整可能。
* **休憩時間**：分単位または時刻区間で入力（実装仕様に合わせて運用）。
* **時給**：任意。入力すると当日・期間集計に賃金概算を表示。
* **保存 / 読み込み**：`work_data.json` への保存、既存データの再読込。
* **集計表示（右ペイン）**：当日・今週・今月の合計勤務時間／（任意）賃金概算。

**メニュー／操作**

* **ファイル**：新規、開く、保存、終了
* **編集**：当日入力のクリア、前日／翌日へ移動
* **表示**：週／月の範囲切替、集計の再計算
* **ヘルプ**：バージョン情報

---

## 💾 データ保存（プライバシー）

* データは既定で **`Working hours calculation application/work_data.json`** に保存されます。
* バックアップ **`work_data.bak`** を生成する場合があります。
* これらは **個人データ** です。公開リポジトリでは **Git 管理対象外** にしてください（`.gitignore` で除外済み）。

> 既にコミットしてしまった場合は `git rm --cached` で追跡解除してください。

---

## 📥 ユーザー向けインストール（.exe 版）

1. [**Releases ページ**](https://github.com/totsuka0405/Working_hours_calculation_application/releases) から最新の実行ファイルをダウンロード
2. 任意のフォルダに `Working_hours_calculation_application.exe` を配置
3. ダブルクリックで起動（インストール不要）

---

## 💻 開発者向け（ソースから実行）

```bash
git clone https://github.com/totsuka0405/Working_hours_calculation_application.git
cd Working_hours_calculation_application

pip install -r requirements.txt
python Working_hours_calculation_application/Working_hours_calculation_application.py
```

---

## 📦 ビルド（.exe 作成：PyInstaller）

### 1) 依存インストール

```bash
pip install pyinstaller
```

### 2) ビルド（アイコン指定の例）

```bash
pyinstaller --noconfirm --onefile --windowed ^
  --icon="Working hours calculation application/work_time_icon.ico" ^
  "Working hours calculation application/Working_hours_calculation_application.py"
```

* 成功すると `dist/Working_hours_calculation_application.exe` が生成されます。

再ビルド前にクリーンアップ推奨：

```bash
rmdir /s /q build dist __pycache__ "Working_hours_calculation_application.spec"
```

---

## 🗂 ディレクトリ構成（抜粋）

```
.
├─ .git/
├─ .gitignore
├─ README.md
├─ Working hours calculation application/
│  ├─ Working_hours_calculation_application.py      # メインスクリプト
│  ├─ Working_hours_calculation_application.pyproj  # VS 用プロジェクト
│  ├─ Working_hours_calculation_application.spec    # PyInstaller 設定
│  ├─ work_time_icon.ico                            # アプリアイコン
│  ├─ work_data.json                                # 個人データ（Git管理外）
│  └─ work_data.bak                                 # 個人データ（Git管理外）
└─ Working hours calculation application.sln        # VS ソリューション
```

---

## 🧪 動作確認チェックリスト

* [ ] 入力（出勤／退勤／休憩）の保存・再読込が正しく行える
* [ ] 当日／週／月の集計値が想定と一致する
* [ ] 時給入力時の概算賃金が計算される
* [ ] アプリ再起動後も前回のデータが復元できる
* [ ] `.exe` 起動がブロックされない（必要に応じて SmartScreen 回避）

---

## ❗ トラブルシューティング

* **exe が起動しない／警告が出る** → SmartScreen の「詳細情報 → 実行」を選択
* **保存先が見つからない** → `work_data.json` はアプリ直下に保存
* **時刻計算の端数処理が違う** → 実装の丸め規則を確認
* **古い exe が残る** → 再ビルド前に `build/ dist/ __pycache__/ *.spec` を削除

---

## 🔒 `.gitignore` 抜粋

```gitignore
# PyInstaller 生成物
build/
dist/
*.spec

# Visual Studio
.vs/
*.suo
*.user

# Python キャッシュ
__pycache__/
*.py[cod]

# 個人データ
Working hours calculation application/work_data.json
Working hours calculation application/work_data.bak
```

---

## 📝 ライセンス

MIT License

---

## 📌 今後の拡張アイデア

* 複数休憩区間の入力対応
* 有給／欠勤など勤務区分の追加
* CSV／Excel 出力
* データ暗号化保存
* 週・月カレンダー表示、祝日対応

---

### リポジトリ

* GitHub: [https://github.com/totsuka0405/Working_hours_calculation_application](https://github.com/totsuka0405/Working_hours_calculation_application)
