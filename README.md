# pub_spushi_to_excel

Google スプレッドシート「test」（testは初期値、設定で変更可能）から、未エクスポート行だけを `.xlsx` で出力し、出力後にエクスポート列へ「済」を書き戻す Apps Script です。

## 最初にここだけ設定
`config.sample.gs` を `config.gs` にコピーして、以下だけ埋めれば動きます。
- `SPREADSHEET_ID`
- `SHEET_NAME`（任意）
- `EXPORT_COLUMN`（任意）
- `OUTPUT_BASENAME`（任意）
- `OUTPUT_FOLDER_ID`（任意）

## セットアップ
1. Apps Script のプロジェクトを用意する（スタンドアロン推奨）。
2. `export_unexported_rows.gs` を追加する。
3. `config.sample.gs` を `config.gs` にコピーして追加・編集する。
   - `SPREADSHEET_ID`: 対象スプレッドシートのID
   - `SHEET_NAME`: `test`（既定値、任意で変更）
   - `EXPORT_COLUMN`: `12`（既定値、任意で変更）
   - `OUTPUT_BASENAME`: `test`（既定値、任意で変更）
   - `OUTPUT_FOLDER_ID`: 空なら元スプレッドシートのフォルダに保存
   - `config.gs` は公開リポジトリに含めない（`.gitignore` 済み）
4. 保存する。

### clasp を使う場合
`clasp clone` でプロジェクトを作り、`clasp push` でファイルを反映できます。
`config.sample.gs` は `.claspignore` により push されません。

## 使い方
- スクリプトエディタから `exportUnexportedRowsToXlsx` を実行する。
- シート名は `SHEET_NAME`（既定値: `test`）、エクスポート列は `EXPORT_COLUMN`（既定値: 12）、エクスポート済みマークは「済」で処理する。
- L 列が空欄の行だけを抽出し、L 列を除外した Excel ファイルを生成する。
- Excel ファイルは元スプレッドシートと同じフォルダ（無ければマイドライブ直下）に作成される。
- 正常終了後、エクスポート対象だった行の L 列に「済」を書き込む。
- 破壊的操作（既存ファイル削除・上書き）は行いません。

## 備考
- 出力ファイル名にはタイムスタンプ（`{ベース名}_yyyyMMdd-HHmmss.xlsx`）が付きます。
- 設定は `config.gs`（`SPREADSHEET_ID`, `SHEET_NAME`, `EXPORT_COLUMN`, `OUTPUT_BASENAME`, `OUTPUT_FOLDER_ID`）または `exportUnexportedRowsToXlsx` 内の `config` で変更できます。
