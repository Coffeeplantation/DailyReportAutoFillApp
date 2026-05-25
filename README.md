# 日報自動入力アプリ

会社から毎月配布される Excel 形式の作業報告書に、勤務時間・備考を自動で入力・保存するアプリです。  
**ウェブ版**（Flask）と **exe 版**（Windows デスクトップ）の 2 種類で利用できます。

---

## 主な機能

- 月〜金の開始・終了・休憩時間を曜日ごとに設定
- 有給取得日をカレンダー UI で選択
- 残業・早退・休日出勤など例外日を個別設定
- 祝日を自動取得し備考欄に「祝日」を記載
- 月間スケジュールのプレビューと合計稼働時間の表示
- 備考欄をプレビュー上で直接編集
- 設定の自動保存（次回起動時に復元）

---

## 自動入力ルール

| 日の種類 | 開始・終了・休憩 | 備考 |
|---|---|---|
| 出勤日（平日） | 曜日別設定の時間 / 休憩 1:00 | 在宅勤務（変更可） |
| 例外日（残業・早退・休日出勤） | 個別設定の時間 | 在宅勤務（変更可） |
| 有給取得日 | 空欄 | 私用により、休暇 |
| 祝日 | 空欄 | 祝日 |
| 土日 | 空欄 | 空欄 |

> 例外日の設定は有給・祝日・土日より最優先で適用されます。

---

## 使い方

### ウェブ版

```bash
pip install flask flask-httpauth openpyxl jpholiday python-dotenv
python app.py
```

ブラウザで `http://localhost:5000/` を開く。  
ログイン情報は `.env` ファイルで設定（デフォルト: `root` / `root`）。

```
# .env
APP_USERNAME=your_username
APP_PASSWORD=your_password
```

### exe 版

[Releases](../../releases/latest) から `DailyReportApp.exe` をダウンロードして実行。  
Python・インストール作業は不要です。

> **初回起動時の注意**  
> Windows Defender SmartScreen の警告が出る場合は「詳細情報」→「実行」を選択してください。

---

## 操作手順

1. 対象年月を確認（自動で当月が設定されます）
2. 曜日別の勤務時間を入力
3. 例外日があれば「＋ 例外日を追加」から設定
4. 有給取得日があればチェックを入れてカレンダーで選択
5. Excel ファイルを選択
6. 「入力完了」ボタンを押す

---

## 設定の保存について

| 項目 | 保存のタイミング |
|---|---|
| 曜日別勤務時間・ラベル | 月をまたいでも引き継ぎ |
| 対象年月・Excelファイルパス | 月をまたいでも引き継ぎ |
| 有給取得日・例外日・手動備考 | 同じ月のみ復元 |

- **ウェブ版**：ブラウザの `localStorage` に自動保存
- **exe 版**：exe と同じフォルダの `settings.json` に自動保存

---

## Excel ファイルの仕様

| 項目 | 内容 |
|---|---|
| 形式 | `.xlsx` |
| ヘッダー行 | 開始時間・終了時間・休憩時間・備考の列名で自動検索 |
| データ開始行 | ヘッダー行の次の行から（月初 1 日） |

列の位置が変わっても、ヘッダーの文字列が一致していれば正常に動作します。

---

## ビルド（exe 版）

GitHub Actions により `main` ブランチへの push 時に自動ビルドされ、Releases に公開されます。

手動でビルドする場合：

```bash
build.bat
```

または：

```bash
pip install pyinstaller openpyxl jpholiday
pyinstaller --onefile --windowed --name "DailyReportApp" daily_report_app.py
```

---

## 技術スタック

| 種別 | 使用技術 |
|---|---|
| ウェブ版 | Python / Flask / openpyxl / jpholiday |
| exe 版 | Python / Tkinter / openpyxl / jpholiday / PyInstaller |
| CI/CD | GitHub Actions |
