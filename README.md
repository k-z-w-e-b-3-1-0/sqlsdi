# sqlsdi

Oracle のテーブルデータを後から再実行できる SQL スクリプトとしてエクスポートするための PowerShell ユーティリティです。ワークフローは条件に一致する行を収集し、対応する `SELECT`、`DELETE`、`INSERT`、および必要に応じてフラッシュバック用のクエリを生成し、結果をタイムスタンプ付きの `.sql` ファイルに保存します。

## リポジトリ構成

| ファイル | 説明 |
| ---- | ----------- |
| `SelectDeletInsert.ps1` | Oracle へ接続してメタデータを取得し、SQL 文を出力するコア関数 `Export-TableSqlWithData`。 |
| `run.ps1` | パラメーターを設定してエクスポート関数を呼び出すサンプル エントリーポイント。 |
| `exec.bat` | `run.ps1` をダブルクリックで実行し、標準出力を `query<timestamp>.sql` にリダイレクトする Windows 用ラッパー。 |

## 必要要件

* PowerShell を備えた Windows 環境（スクリプトでは `C:\\Windows\\SysWOW64\\WindowsPowerShell\\v1.0` 配下の `powershell.exe` を想定）。
* Oracle Data Provider for .NET (`Oracle.DataAccess.dll`)。アセンブリがグローバル アセンブリ キャッシュに存在しない場合は、`ORACLE_DLL_PATH` 環境変数を設定するか、`run.ps1` 実行時に `-OracleDllPath` パラメーターを指定します。
* 指定した接続文字列で対象の Oracle データベースに接続できるネットワーク環境。

## 使い方

1. Oracle クライアントがインストールされた Windows マシン上の任意のフォルダーにスクリプト一式をコピーします。
2. `run.ps1` 内のプレースホルダーを編集します。
   * `-ConnectionString` と `-Schema` に渡している `USERID`、`PASSWORD`、`IPADDRESS:PORT/SID`、`SCHEMA` を実環境に合わせて置き換えます。
   * 対象列（多くの場合はステータス列やタイムスタンプ列）と SQL 条件を設定します。`exec.bat` に指定した値が `$Column` と `$Condition` として `run.ps1` に渡されます。
   * 任意: データベースがフラッシュバックをサポートしない場合は `-FlashbackTimestamp` を削除するか、必要なタイムスタンプに調整します。
   * 任意: データエクスポートの前処理として実行したい PL/SQL ブロックがあれば `$plsql` 変数を変更します。スクリプトは `:result` 出力パラメーターをログに追記します。
3. `sql/primary_keys.sql` と `sql/column_definitions.sql` のテンプレート SQL を上書きするか、パラメーターで別パスを指定することもできます（任意）。
4. `exec.bat` を実行します。バッチ ファイルは作業ディレクトリの設定、タイムスタンプの生成、PowerShell 経由での `run.ps1` 実行を行います。生成された SQL は `output\\query_<timestamp>.sql` に書き出され、同時にコンソールにも出力されます。処理が完了したら任意のキーを押してウィンドウを閉じてください。

### バッチ ファイルを使用しない場合

PowerShell から直接スクリプトを呼び出すこともできます。

```powershell
# リポジトリ フォルダー内の PowerShell から
. .\SelectDeletInsert.ps1
Export-TableSqlWithData -ConnectionString "User Id=..." -Schema "SCHEMA" -TargetColumn "column_name" -ConditionSql "= 'value'" -FlashbackTimestamp "2025-08-17 19:00:00"
```

スクリプトは 4 つのセクション（`SELECT`、`DELETE`、`INSERT`、任意で `FLASHBACK`）を順番に出力し、最後に警告などを記録するログ領域が続きます。`exec.bat` を使用しない場合は、出力をファイルにリダイレクトしてください。

```powershell
Export-TableSqlWithData ... | Out-File query.sql -Encoding UTF8
```

## 注意事項

* エクスポート対象のテーブルには必ず主キーが存在する必要があります。関数は主キーのメタデータを利用して行の並びと出力を決定します。
* 文字列はシングルクォートを二重化してエスケープし、Oracle の日付／タイムスタンプは `TO_DATE` を利用してフォーマットします。
* `INSERT` 文では Null 値をリテラルの `NULL` として出力します。
* 指定スキーマ内で対象列を持つテーブルのみを抽出し、列コメントを取得して出力の見出しとして付与します。
