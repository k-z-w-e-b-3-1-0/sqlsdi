# sqlsdi

PowerShell utilities for exporting Oracle table data as SQL scripts that can be replayed later. The workflow collects matching rows, generates matching `SELECT`, `DELETE`, `INSERT`, and optional flashback queries, and saves the output to a timestamped `.sql` file.

## Repository layout

| File | Description |
| ---- | ----------- |
| `SelectDeletInsert.ps1` | Core function `Export-TableSqlWithData` that connects to Oracle, retrieves metadata, and emits SQL statements. |
| `run.ps1` | Example entry point that wires parameters and calls the export function. |
| `exec.bat` | Convenience wrapper for Windows to invoke `run.ps1` from a double-click and redirect the output to `query<timestamp>.sql`. |

## Requirements

* Windows with PowerShell (the scripts assume `powershell.exe` under `C:\Windows\SysWOW64\WindowsPowerShell\v1.0`).
* Oracle Data Provider for .NET (`Oracle.DataAccess.dll`). Update the hard-coded path inside `SelectDeletInsert.ps1` to point to the DLL on your machine before running the scripts.
* Database connectivity that allows the supplied connection string to reach the target Oracle database.

## Usage

1. Copy the scripts to a folder on a Windows machine where the Oracle client is installed.
2. Edit the placeholders in `run.ps1`:
   * Replace `USERID`, `PASSWORD`, `IPADDRESS:PORT/SID`, and `SCHEMA` in the `-ConnectionString` and `-Schema` arguments.
   * Supply a target column name (typically a status or timestamp column) and SQL condition. The values passed through `exec.bat` become the `$Column` and `$Condition` parameters consumed by `run.ps1`.
   * Optional: adjust the flashback timestamp or remove the `-FlashbackTimestamp` argument if the database does not support flashback queries.
   * Optional: modify `$plsql` with any pre-processing block that should execute before the data export. The script captures the `:result` output parameter and appends the message to the log section.
3. (Optional) Change the `Oracle.DataAccess.dll` path in `SelectDeletInsert.ps1` and confirm it matches your environment.
4. Run `exec.bat`. The batch file sets the working directory, computes a timestamp, and runs `run.ps1` through PowerShell, saving the generated SQL to `query<timestamp>.sql` in the same directory. When execution finishes, press any key to close the window.

### Running without the batch file

You can call the PowerShell script directly:

```powershell
# From a PowerShell prompt inside the repository folder
. .\SelectDeletInsert.ps1
Export-TableSqlWithData -ConnectionString "User Id=..." -Schema "SCHEMA" -TargetColumn "column_name" -ConditionSql "= 'value'" -FlashbackTimestamp "2025-08-17 19:00:00"
```

The script emits four sections (`SELECT`, `DELETE`, `INSERT`, and optional `FLASHBACK`) followed by a log area with any warnings encountered. Redirect the output to a file if you do not use `exec.bat`:

```powershell
Export-TableSqlWithData ... | Out-File query.sql -Encoding UTF8
```

## Notes

* Ensure that every table participating in the export has a primary key. The function uses the key metadata to order rows and build deterministic output.
* Character values are escaped with doubled single quotes, and Oracle date/timestamp values are formatted via `TO_DATE`.
* Null values are emitted as literal `NULL` in the generated `INSERT` statements.
* The script filters tables that contain the target column and belong to the specified schema, pulling column comments to include them as captions in the output.
