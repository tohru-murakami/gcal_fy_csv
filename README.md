# Googleカレンダー → CSV変換ツール

このツールは Googleカレンダーの予定をCSVに変換します。
今年度（4月1日〜翌年3月31日）の予定だけを出力します。

出力ファイル：

`gunma_FY.csv`

Excelで開くことができます。

##  ICS URLの取得方法

Googleカレンダーを開きます：

`https://calendar.google.com/`

左側の「マイカレンダー」で対象のカレンダーの
「︙」をクリックします。

「設定と共有」をクリックします。


画面を下にスクロールして

「カレンダーの統合」

を探します。


次の項目を見つけます：

秘密のアドレス（iCal形式）


表示されるURLをコピーします。


URLの例：`https://calendar.google.com/calendar/ical/XXXXXXXX%40group.calendar.google.com/private-XXXXXXXX/basic.ics`

非公開のカレンダーのURLは秘密情報です。他人に公開しないでください。

## Macでの使い方

1. mac_gcal_fy_csv.command を開きます

2. 次の行を探します：

   `ICS_URL="https://calendar.google.com/calendar/ical/XXXX/basic.ics"`

3. XXXX の部分をコピーしたURLに置き換えます

   例：`ICS_URL="https://calendar.google.com/calendar/ical/XXXXXXXX/basic.ics"`

4. ファイルを保存します

5. 初回だけ

   `chmod +x mac_gcal_fy_csv.command`

6. ダブルクリックします

   CSVファイルが作成されます：

   `gunma_FY.csv`

## Windowsでの使い方

1. `win_gcal_fy_csv.ps1` をメモ帳で開きます

2. 次の行を探します：

   `$IcsUrl = "https://calendar.google.com/calendar/ical/XXXX/basic.ics"`

3. XXXX の部分をコピーしたURLに置き換えます

4. 保存します

5. PowerShellを開きます

6. このフォルダに移動します

   例：`cd Desktop\calendar_export`

7. 実行します：

   `Set-ExecutionPolicy -Scope Process Bypass`

   `.\win_gcal_fy_csv.ps1`

8. CSVファイルが作成されます：

   `gunma_FY.csv`

## トラブル対処

- CSVが空の場合

  - URLが間違っている可能性があります。

- Excelで文字化けする場合

  - Excelの「データ → テキストから」で読み込んでください。

- Macで開けない場合

  - 最初に一度だけ実行します：

    `chmod +x mac_gcal_fy_csv.command`

## 注意

ICS URLは秘密情報です。

URLを知っている人は予定表を取得できます。

Webページやメールに公開しないでください。