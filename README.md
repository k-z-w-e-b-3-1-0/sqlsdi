# sqlsdi

Oracle のテーブルデータを後から再実行できる SQL スクリプトとしてエクスポートするための PowerShell ユーティリティです。ワークフローは条件に一致する行を収集し、対応する `SELECT`、`DELETE`、`INSERT`、および必要に応じてフラッシュバック用のクエリを生成し、結果をタイムスタンプ付きの `.sql` ファイルに保存します。

## リポジトリ構成

| ファイル | 説明 |
| ---- | ----------- |
| `SelectDeletInsert/SelectDeletInsert.ps1` | Oracle へ接続してメタデータを取得し、SQL 文を出力するコア関数 `Export-TableSqlWithData`。 |
| `SelectDeletInsert/run.ps1` | 単一列条件用のエクスポート関数を呼び出すサンプル エントリーポイント。 |
| `SelectDeletInsertIncrement/SelectDeletInsertIncrement.ps1` | 取得したデータに加算処理を適用して出力する `Export-TableSqlWithIncrementedData`。 |
| `SelectDeletInsertIncrement/run.ps1` | インクリメント用の関数に対して対象列や加算値を渡すエントリーポイント。 |
| `SelectDeletInsertMulti/SelectDeletInsertMulti.ps1` | 複数列を対象にした条件組み立てとエクスポートを行う `Export-TableSqlWithDataMultiColumn`。 |
| `SelectDeletInsertMulti/run.ps1` | 列配列や列ごとの条件マップを受け取り、マルチカラム抽出を実行するエントリーポイント。 |
| `exec.bat` | `SelectDeletInsert\run.ps1` をダブルクリックで実行し、標準出力を `query<timestamp>.sql` にリダイレクトする Windows 用ラッパー。 |

## 必要要件

* PowerShell を備えた Windows 環境（スクリプトでは `C:\\Windows\\SysWOW64\\WindowsPowerShell\\v1.0` 配下の `powershell.exe` を想定）。
* Oracle Data Provider for .NET (`Oracle.DataAccess.dll`)。アセンブリがグローバル アセンブリ キャッシュに存在しない場合は、`ORACLE_DLL_PATH` 環境変数を設定するか、各 `run.ps1`（例: `SelectDeletInsert/run.ps1`）実行時に `-OracleDllPath` パラメーターを指定します。
* 指定した接続文字列で対象の Oracle データベースに接続できるネットワーク環境。

## 使い方

1. Oracle クライアントがインストールされた Windows マシン上の任意のフォルダーにスクリプト一式をコピーします。
2. `SelectDeletInsert/run.ps1`（単一列向け）、`SelectDeletInsertIncrement/run.ps1`（加算付き単一列向け）、`SelectDeletInsertMulti/run.ps1`（複数列向け）内のプレースホルダーを編集します。
   * `-ConnectionString` と `-Schema` に渡している `USERID`、`PASSWORD`、`IPADDRESS:PORT/SID`、`SCHEMA` を実環境に合わせて置き換えます。
   * 単一列の場合は対象列と SQL 条件を設定します。`exec.bat` に指定した値が `$Column` と `$Condition` として `SelectDeletInsert/run.ps1` に渡されます。
   * 複数列で抽出する場合は `SelectDeletInsertMulti/run.ps1` の `$Columns` 配列、`$ConditionFragments`、`$EqualConditions` のいずれか（もしくは組み合わせ）を編集します。列名をキーにした等価条件マップを用意すると、各列に `=` 条件を一括で適用できます。
   * インクリメント付きの抽出を行う場合は `SelectDeletInsertIncrement/run.ps1` の `$IncrementColumns` や `$IncrementOverrides` を調整し、上書きしたい列と加算値を指定します。
   * `-OutputDirectory` パラメーターにエクスポート先のフォルダーを指定できます。未指定の場合は各 `run.ps1` と同じディレクトリ配下の `output` フォルダーが利用されます。
   * 任意: データベースがフラッシュバックをサポートしない場合は `-FlashbackTimestamp` を削除するか、必要なタイムスタンプに調整します。
   * 任意: データエクスポートの前処理として実行したい PL/SQL ブロックがあれば `$plsql` 変数を変更します。スクリプトは `:result` 出力パラメーターをログに追記します。
3. `sql/primary_keys.sql` と `sql/column_definitions.sql` のテンプレート SQL を上書きするか、パラメーターで別パスを指定することもできます（任意）。
4. `exec.bat` を実行します。バッチ ファイルは作業ディレクトリの設定、タイムスタンプの生成、PowerShell 経由での `SelectDeletInsert\run.ps1` 実行を行います。生成された SQL は `output\\query_<timestamp>.sql` に書き出され、同時にコンソールにも出力されます。処理が完了したら任意のキーを押してウィンドウを閉じてください。

### バッチ ファイルを使用しない場合

PowerShell から直接スクリプトを呼び出すこともできます。

```powershell
# リポジトリ フォルダー内の PowerShell から
. .\SelectDeletInsert\SelectDeletInsert.ps1
Export-TableSqlWithData -ConnectionString "User Id=..." -Schema "SCHEMA" -TargetColumn "column_name" -ConditionSql "= 'value'" -FlashbackTimestamp "2025-08-17 19:00:00"
```

複数列を条件にする場合はマルチカラム用の関数を読み込みます。

```powershell
. .\SelectDeletInsertMulti\SelectDeletInsertMulti.ps1
Export-TableSqlWithDataMultiColumn -ConnectionString "User Id=..." -Schema "SCHEMA" -TargetColumns @('PK_A', 'PK_B') -ConditionFragments @("PK_A = '001'", "PK_B = 'XYZ'")
```

スクリプトは 4 つのセクション（`SELECT`、`DELETE`、`INSERT`、任意で `FLASHBACK`）を順番に出力し、最後に警告などを記録するログ領域が続きます。`exec.bat` を使用しない場合は、出力をファイルにリダイレクトしてください。

```powershell
Export-TableSqlWithData ... | Out-File query.sql -Encoding UTF8
```

### インクリメント付きエクスポート スクリプトの実行例

列値に加算処理を適用したい場合は `SelectDeletInsertIncrement/SelectDeletInsertIncrement.ps1` を読み込み、`Export-TableSqlWithIncrementedData` を呼び出します。`-IncrementColumns` で対象列を指定し、`-IncrementValue` で一括加算値を、列ごとに異なる加算値を指定したい場合は `-IncrementOverrides` にハッシュテーブルを渡します。

```powershell
. .\SelectDeletInsertIncrement\SelectDeletInsertIncrement.ps1
Export-TableSqlWithIncrementedData \ 
  -ConnectionString "User Id=..." \ 
  -Schema "SCHEMA" \ 
  -TargetColumn "COL_A" \ 
  -ConditionSql "BETWEEN '2024-01-01' AND '2024-01-31'" \ 
  -IncrementColumns @('AMOUNT', 'TAX') \ 
  -IncrementValue 1.5 \ 
  -IncrementOverrides @{ 'TAX' = 0.8 } \ 
  -OutputFile ".\\output\\query_increment.sql"
```

上記の例では `AMOUNT` 列には 1.5、`TAX` 列には 0.8 が加算された状態で `SELECT`、`DELETE`、`INSERT` の各セクションが出力されます。

## Excel からの実行例 (VBA モジュール)

Excel の一覧から `SelectDeletInsert/run.ps1` を順番に起動したい場合は、リポジトリの `excel/RunPowerShell.bas` モジュールをインポートしてください。

1. Excel で [開発] タブ → [Visual Basic] を開き、`ファイル` → `ファイルのインポート` から `excel/RunPowerShell.bas` を読み込みます。
2. 既定では `Sheet1` の A～E 列を下表のように解釈します。1 行目はヘッダー、2 行目以降をデータ行として処理します。

   | 列 | 項目 | 説明 |
   | --- | ---- | ---- |
   | A | `TargetColumn` | `SelectDeletInsert/run.ps1` の `-Column` に渡す列名。必須。 |
   | B | `ConditionSql` | `SelectDeletInsert/run.ps1` の `-Condition` に渡す WHERE 条件。必須。 |
   | C | `ConnectionString` | 行ごとの接続文字列。空欄の場合は後述の名前付きセル `DefaultConnectionString` を参照します。 |
   | D | `OracleDllPath` | `-OracleDllPath` に渡す DLL パス。空欄の場合は `DefaultOracleDllPath` の値を利用します。 |
   | E | `OutputDirectory` | 出力フォルダー。空欄の場合は `DefaultOutputDirectory` の値、さらに未設定なら `SelectDeletInsert/run.ps1` の既定値 (`output` フォルダー) を使用します。 |

3. `SelectDeletInsert/run.ps1` までのパスは、ブックと同じフォルダーに配置した場合は自動で解決されます。別の場所に置く場合は、任意のセルにフルパスまたは相対パスを入力し、そのセルに名前 `PowerShellScriptPath` を定義してください。相対パスはブックの保存先を起点に解決され、`%USERPROFILE%` などの環境変数も利用できます。
4. 行単位で既定値を補いたい場合は、以下の名前付きセルを定義します (必要なものだけで構いません)。

   * `DefaultConnectionString`
   * `DefaultOracleDllPath`
   * `DefaultOutputDirectory`

   それぞれ空欄の列に対するフォールバックとして使用されます。

5. モジュールには次のマクロが含まれています。

   * `RunPowerShellBatch` – 一覧の全行を順次実行し、各 PowerShell プロセスの完了を待ちます。
   * `RunPowerShellBatchAsync` – 一覧を順次起動しますが、各行の完了を待たずに次へ進みます。
   * `RunPowerShellForActiveRow` – 現在選択している行だけを実行します。

   いずれのマクロも `powershell.exe -NoProfile -ExecutionPolicy Bypass` で `SelectDeletInsert/run.ps1` を起動し、列の値に応じて `-ConnectionString`、`-OracleDllPath`、`-OutputDirectory` を自動的に付与します。

> **ヒント:** Excel から呼び出す前に PowerShell 単体で `SelectDeletInsert/run.ps1` が期待通りに動作することを確認し、実行ポリシー (`Set-ExecutionPolicy`) やマクロ実行権限を適切に設定してください。

## 注意事項

* エクスポート対象のテーブルには必ず主キーが存在する必要があります。関数は主キーのメタデータを利用して行の並びと出力を決定します。
* 文字列はシングルクォートを二重化してエスケープし、Oracle の日付／タイムスタンプは `TO_DATE` を利用してフォーマットします。
* `INSERT` 文では Null 値をリテラルの `NULL` として出力します。
* 指定スキーマ内で対象列を持つテーブルのみを抽出し、列コメントを取得して出力の見出しとして付与します。
