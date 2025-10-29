function Export-TableSqlWithData {
    param(
        [string]$ConnectionString,
        [string]$Schema,
        [string]$TargetColumn,
        [string]$ConditionSql,
        [string]$FlashbackTimestamp = $null,
        [string]$PreSql = $null,
        [string]$OracleDllPath = $null,
        [string]$PrimaryKeysSqlPath = $null,
        [string]$ColumnDefinitionsSqlPath = $null,
        [string]$OutputFile = $null
    )

    if ($OracleDllPath) {
        if (-not (Test-Path -Path $OracleDllPath)) {
            throw "Oracle Data Provider DLL not found at '$OracleDllPath'."
        }
        [void][Reflection.Assembly]::LoadFrom($OracleDllPath)
    } else {
        try {
            [void][Reflection.Assembly]::Load("Oracle.DataAccess")
        } catch {
            throw "Unable to load Oracle.DataAccess. Specify -OracleDllPath with the full path to Oracle.DataAccess.dll."
        }
    }

    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptDirectory) {
        $scriptDirectory = Get-Location
    }

    $defaultPrimaryKeySql = Join-Path $scriptDirectory "sql/primary_keys.sql"
    $defaultColumnDefinitionsSql = Join-Path $scriptDirectory "sql/column_definitions.sql"

    $primaryKeySqlPath = if ($PrimaryKeysSqlPath) { $PrimaryKeysSqlPath } else { $defaultPrimaryKeySql }
    $columnDefinitionsSqlPath = if ($ColumnDefinitionsSqlPath) { $ColumnDefinitionsSqlPath } else { $defaultColumnDefinitionsSql }

    if (-not (Test-Path -Path $primaryKeySqlPath)) {
        throw "Primary key SQL file not found at '$primaryKeySqlPath'."
    }
    if (-not (Test-Path -Path $columnDefinitionsSqlPath)) {
        throw "Column definitions SQL file not found at '$columnDefinitionsSqlPath'."
    }

    $primaryKeysSql = Get-Content -Path $primaryKeySqlPath -Raw
    $columnDefinitionsSql = Get-Content -Path $columnDefinitionsSqlPath -Raw

    $outputLines = New-Object System.Collections.Generic.List[string]

    $conn = New-Object Oracle.DataAccess.Client.OracleConnection($ConnectionString)
    $conn.Open()

    $selectLines     = @()
    $deleteLines     = @()
    $insertLines     = @()
    $flashbackLines  = @()
    $logLines        = @()

    try {
        # 任意の PL/SQL 実行
        if ($PreSql) {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $PreSql
            $cmd.CommandType = [System.Data.CommandType]::Text
            $outParam = $cmd.Parameters.Add("result", [Oracle.DataAccess.Client.OracleDbType]::Varchar2, 4000)
            $outParam.Direction = [System.Data.ParameterDirection]::Output
            try {
                $cmd.ExecuteNonQuery() | Out-Null
                $logLines += "-- PreSql result: $($outParam.Value)"
            } catch {
                $logLines += "-- PreSql execution failed: $_"
            }
        }

        # 主キー列取得
        $primaryKeys = @{}
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $primaryKeysSql
        $cmd.CommandType = [System.Data.CommandType]::Text
        $cmd.BindByName = $true
        $schemaParam = $cmd.Parameters.Add("schema", [Oracle.DataAccess.Client.OracleDbType]::Varchar2, 128)
        $schemaParam.Value = $Schema

        $reader = $cmd.ExecuteReader()
        try {
            while ($reader.Read()) {
                $owner = $reader.GetString(0)
                $table = $reader.GetString(1)
                $column = $reader.GetString(2)
                $key = "$table"
                if (-not $primaryKeys.ContainsKey($key)) {
                    $primaryKeys[$key] = @()
                }
                $primaryKeys[$key] += $column
            }
        } finally {
            $reader.Close()
        }

        # 列定義取得
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $columnDefinitionsSql
        $cmd.CommandType = [System.Data.CommandType]::Text
        $cmd.BindByName = $true
        $schemaParam = $cmd.Parameters.Add("schema", [Oracle.DataAccess.Client.OracleDbType]::Varchar2, 128)
        $schemaParam.Value = $Schema
        $targetColumnParam = $cmd.Parameters.Add("target_column", [Oracle.DataAccess.Client.OracleDbType]::Varchar2, 128)
        $targetColumnParam.Value = $TargetColumn

        $tables = @{}
        $reader = $cmd.ExecuteReader()
        try {
            while ($reader.Read()) {
                $owner = $reader.GetString(0)
                $table = $reader.GetString(1)
                $column = $reader.GetString(2)
                $comment = if (!$reader.IsDBNull(3)) {
                    $c = $reader.GetString(3)
                    if ([string]::IsNullOrEmpty($c)) { "" } else { $c }
                } else { "" }
                $dataType = $reader.GetString(5)
                $key = "$table"
                if (-not $tables.ContainsKey($key)) { $tables[$key] = @() }
                $tables[$key] += ,@($column, $comment, $dataType)
            }
        } finally {
            $reader.Close()
        }

        # テーブルごとにデータ抽出
        $counter = 1
        foreach ($key in $tables.Keys  | Sort-Object) {
            $meta = $tables[$key]
            $colNames = $meta | ForEach-Object { $_[0] }
            $flatCols = $colNames
            $flatComments = $meta | ForEach-Object { $_[1] }
            $whereClause = "$TargetColumn $ConditionSql"
            $orderClause = ""
            if ($primaryKeys.ContainsKey($key) -and $primaryKeys[$key].Count -gt 0) {
                $orderClause = " ORDER BY " + ($primaryKeys[$key] -join ", ")
            }

            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT " + ($flatCols -join ", ") + " FROM $key WHERE $whereClause"
            $cmd.CommandType = [System.Data.CommandType]::Text

            $reader = $cmd.ExecuteReader()

            try {
                if ($reader.HasRows) {

                    $selectLines += "-- $counter $key"
                    $selectLines += "-- * Captions " + ($flatComments -join ", ")
                    $selectSql = "SELECT " + ($flatCols -join ", ") + " FROM $key WHERE $whereClause" + $orderClause + ";"
                    $selectLines += $selectSql

                    $deleteLines += "DELETE FROM $key WHERE $whereClause;"

                    while ($reader.Read()) {
                        $vals = @()
                        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                            $v = $reader.GetValue($i)
                            $t = $meta[$i][2]
                            if ($v -eq $null -or $reader.IsDBNull($i)) {
                                $vals += "NULL"
                            } elseif ($t -like "*CHAR*" -or $t -like "*CLOB*") {
                                $vals += "'" + $v.ToString().Replace("'", "''") + "'"
                            } elseif ($t -like "*DATE*" -or $t -like "*TIMESTAMP*") {
                                try {
                                    $dt = $reader.GetDateTime($i)
                                    $vals += "TO_DATE('" + $dt.ToString("yyyy-MM-dd HH:mm:ss") + "','YYYY-MM-DD HH24:MI:SS')"
                                } catch {
                                    $logLines += "-- 日付変換失敗: $($_.Exception.Message)"
                                    $vals += "NULL"
                                }
                            } else {
                                $vals += $v.ToString()
                            }
                        }
                        $insertLines += "INSERT INTO $key (" + ($flatCols -join ", ") + ") VALUES (" + ($vals -join ", ") + ");"
                    }

                    if ($FlashbackTimestamp) {
                        $flashbackLines += "-- $counter $key"
                        $fbSql = "SELECT " + ($flatCols -join ", ") + " FROM $key AS OF TIMESTAMP TO_TIMESTAMP('$FlashbackTimestamp','YYYY-MM-DD HH24:MI:SS') WHERE $whereClause"
                        if ($orderClause) { $fbSql += $orderClause }
                        $fbSql += ";"
                        $flashbackLines += "-- * Captions " + ($flatComments -join ", ")
                        $flashbackLines += $fbSql
                    }

                    $counter++
                }
            } finally {
                $reader.Close()
            }
        }
    } finally {
        $conn.Close()
    }

    $sections = @(
        @{ Header = "SELECT"; Lines = $selectLines },
        @{ Header = "DELETE"; Lines = $deleteLines },
        @{ Header = "INSERT"; Lines = $insertLines }
    )

    if ($flashbackLines.Count -gt 0) {
        $sections += @{ Header = "FLASHBACK"; Lines = $flashbackLines }
    }
    if ($logLines.Count -gt 0) {
        $sections += @{ Header = "LOG / WARNINGS"; Lines = $logLines }
    }

    foreach ($section in $sections) {
        $headerLine = "`n-------- $($section.Header) --------"
        Write-Output $headerLine
        $outputLines.Add($headerLine) | Out-Null
        foreach ($line in $section.Lines) {
            Write-Output $line
            $outputLines.Add($line) | Out-Null
        }
    }

    if ($OutputFile) {
        $outputDirectory = Split-Path -Parent $OutputFile
        if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        $outputLines | Set-Content -Path $OutputFile -Encoding UTF8
    }
}
