if (-not (Get-Command Get-SafeFileName -ErrorAction SilentlyContinue)) {
    function Get-SafeFileName {
        param(
            [string]$Value
        )

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return ""
        }

        $safe = $Value -replace '\s+', '_'
        $safe = $safe -replace '[^0-9A-Za-z_\-]', '_'
        $safe = [System.Text.RegularExpressions.Regex]::Replace($safe, '_{2,}', '_')
        $safe = $safe.Trim('_')

        if ($safe.Length -gt 80) {
            $safe = $safe.Substring(0, 80)
        }

        return $safe
    }
}

function Resolve-SqlDirectory {
    param(
        [string]$StartDirectory
    )

    $checkedDirectories = @()
    $current = $StartDirectory

    while ($current -and ($checkedDirectories -notcontains $current)) {
        $checkedDirectories += $current
        $sqlCandidate = Join-Path $current 'sql'
        if (Test-Path -Path $sqlCandidate) {
            return $sqlCandidate
        }

        $current = Split-Path -Parent $current
        if (-not $current) {
            break
        }
    }

    throw "SQL directory not found relative to '$StartDirectory'."
}

function ConvertToSqlLiteral {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return "NULL"
    }

    if ($Value -is [System.DateTime]) {
        $dt = [System.DateTime]$Value
        return "TO_DATE('" + $dt.ToString('yyyy-MM-dd HH:mm:ss') + "','YYYY-MM-DD HH24:MI:SS')"
    }

    if ($Value -is [System.Boolean]) {
        return if ($Value) { "1" } else { "0" }
    }

    if ($Value -is [System.Byte] -or
        $Value -is [System.SByte] -or
        $Value -is [System.Int16] -or
        $Value -is [System.UInt16] -or
        $Value -is [System.Int32] -or
        $Value -is [System.UInt32] -or
        $Value -is [System.Int64] -or
        $Value -is [System.UInt64] -or
        $Value -is [System.Single] -or
        $Value -is [System.Double] -or
        $Value -is [System.Decimal]) {
        return [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
    }

    $escaped = $Value.ToString().Replace("'", "''")
    return "'${escaped}'"
}

function Export-TableSqlWithDataMultiColumn {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConnectionString,

        [Parameter(Mandatory = $true)]
        [string]$Schema,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetColumns,

        [string]$ConditionSql,

        [string[]]$ConditionFragments,

        [hashtable]$EqualityConditions,

        [string]$FlashbackTimestamp = $null,

        [string]$PreSql = $null,

        [string]$OracleDllPath = $null,

        [string]$PrimaryKeysSqlPath = $null,

        [string]$OutputFile = $null,

        [switch]$SplitByOperation
    )

    if (-not $TargetColumns -or $TargetColumns.Count -eq 0) {
        throw "TargetColumns must include at least one column name."
    }

    $normalizedColumns = @()
    foreach ($column in $TargetColumns) {
        if ([string]::IsNullOrWhiteSpace($column)) { continue }
        $normalizedColumns += $column.Trim()
    }

    if ($normalizedColumns.Count -eq 0) {
        throw "TargetColumns must include at least one non-empty column name."
    }

    if (-not $ConditionSql) {
        $clauses = New-Object System.Collections.Generic.List[string]

        if ($ConditionFragments) {
            foreach ($fragment in $ConditionFragments) {
                if (-not [string]::IsNullOrWhiteSpace($fragment)) {
                    $clauses.Add($fragment.Trim()) | Out-Null
                }
            }
        }

        if ($EqualityConditions) {
            foreach ($entry in $EqualityConditions.GetEnumerator()) {
                $columnName = $entry.Name
                $value = $entry.Value
                if ([string]::IsNullOrWhiteSpace($columnName)) { continue }
                $literal = ConvertToSqlLiteral -Value $value
                if ($literal -eq 'NULL') {
                    $clauses.Add("$columnName IS NULL") | Out-Null
                } else {
                    $clauses.Add("$columnName = $literal") | Out-Null
                }
            }
        }

        if ($clauses.Count -eq 0) {
            throw "ConditionSql is empty. Provide ConditionSql, ConditionFragments, or EqualityConditions."
        }

        $ConditionSql = ($clauses -join ' AND ')
    }

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

    $sqlDirectory = Resolve-SqlDirectory -StartDirectory $scriptDirectory
    $defaultPrimaryKeySql = Join-Path $sqlDirectory "primary_keys.sql"
    $primaryKeysSqlPath = if ($PrimaryKeysSqlPath) { $PrimaryKeysSqlPath } else { $defaultPrimaryKeySql }

    if (-not (Test-Path -Path $primaryKeysSqlPath)) {
        throw "Primary key SQL file not found at '$primaryKeysSqlPath'."
    }

    $primaryKeysSql = Get-Content -Path $primaryKeysSqlPath -Raw

    $outputLines = New-Object System.Collections.Generic.List[string]

    $conn = New-Object Oracle.DataAccess.Client.OracleConnection($ConnectionString)
    $conn.Open()

    $selectLines     = @()
    $deleteLines     = @()
    $insertLines     = @()
    $flashbackLines  = @()
    $logLines        = @()

    try {
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
                $table = $reader.GetString(1)
                $column = $reader.GetString(2)
                if (-not $primaryKeys.ContainsKey($table)) {
                    $primaryKeys[$table] = @()
                }
                $primaryKeys[$table] += $column
            }
        } finally {
            $reader.Close()
        }

        $columnQueryBuilder = New-Object System.Text.StringBuilder
        [void]$columnQueryBuilder.AppendLine('SELECT :schema AS OWNER,')
        [void]$columnQueryBuilder.AppendLine('       c.TABLE_NAME,')
        [void]$columnQueryBuilder.AppendLine('       c.COLUMN_NAME,')
        [void]$columnQueryBuilder.AppendLine('       com.COMMENTS,')
        [void]$columnQueryBuilder.AppendLine('       c.COLUMN_ID,')
        [void]$columnQueryBuilder.AppendLine('       c.DATA_TYPE')
        [void]$columnQueryBuilder.AppendLine('  FROM USER_TAB_COLUMNS c')
        [void]$columnQueryBuilder.AppendLine('  LEFT JOIN USER_COL_COMMENTS com')
        [void]$columnQueryBuilder.AppendLine('         ON c.TABLE_NAME = com.TABLE_NAME')
        [void]$columnQueryBuilder.AppendLine('        AND c.COLUMN_NAME = com.COLUMN_NAME')
        [void]$columnQueryBuilder.AppendLine(' WHERE EXISTS (SELECT 1 FROM USER_TABLES u WHERE u.TABLE_NAME = c.TABLE_NAME)')

        $index = 0
        foreach ($column in $normalizedColumns) {
            $placeholder = "target_column_$index"
            [void]$columnQueryBuilder.AppendLine("   AND EXISTS (SELECT 1 FROM USER_TAB_COLUMNS t WHERE t.TABLE_NAME = c.TABLE_NAME AND t.COLUMN_NAME = :$placeholder)")
            $index++
        }

        [void]$columnQueryBuilder.AppendLine(' ORDER BY c.TABLE_NAME, c.COLUMN_ID')

        $columnSql = $columnQueryBuilder.ToString()

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $columnSql
        $cmd.CommandType = [System.Data.CommandType]::Text
        $cmd.BindByName = $true
        $schemaParam = $cmd.Parameters.Add("schema", [Oracle.DataAccess.Client.OracleDbType]::Varchar2, 128)
        $schemaParam.Value = $Schema

        $index = 0
        foreach ($column in $normalizedColumns) {
            $param = $cmd.Parameters.Add("target_column_$index", [Oracle.DataAccess.Client.OracleDbType]::Varchar2, 128)
            $param.Value = $column
            $index++
        }

        $tables = @{}
        $reader = $cmd.ExecuteReader()
        try {
            while ($reader.Read()) {
                $table = $reader.GetString(1)
                $column = $reader.GetString(2)
                $comment = if (!$reader.IsDBNull(3)) {
                    $value = $reader.GetString(3)
                    if ([string]::IsNullOrEmpty($value)) { "" } else { $value }
                } else { "" }
                $dataType = $reader.GetString(5)
                if (-not $tables.ContainsKey($table)) { $tables[$table] = @() }
                $tables[$table] += ,@($column, $comment, $dataType)
            }
        } finally {
            $reader.Close()
        }

        $counter = 1
        foreach ($table in ($tables.Keys | Sort-Object)) {
            $meta = $tables[$table]
            $columnNames = $meta | ForEach-Object { $_[0] }
            $columnComments = $meta | ForEach-Object { $_[1] }
            $dataTypes = $meta | ForEach-Object { $_[2] }

            $whereClause = $ConditionSql
            $orderClause = ""
            if ($primaryKeys.ContainsKey($table) -and $primaryKeys[$table].Count -gt 0) {
                $orderClause = " ORDER BY " + ($primaryKeys[$table] -join ', ')
            }

            $selectSql = "SELECT " + ($columnNames -join ', ') + " FROM $table WHERE $whereClause"
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $selectSql
            $cmd.CommandType = [System.Data.CommandType]::Text

            $reader = $cmd.ExecuteReader()
            try {
                if ($reader.HasRows) {
                    $selectLines += "-- $counter $table"
                    $selectLines += "-- * Captions " + ($columnComments -join ', ')
                    $selectLines += $selectSql + $orderClause + ';'

                    $deleteLines += "DELETE FROM $table WHERE $whereClause;"

                    while ($reader.Read()) {
                        $values = @()
                        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                            $value = $reader.GetValue($i)
                            $dataType = $dataTypes[$i]
                            if ($reader.IsDBNull($i) -or $null -eq $value) {
                                $values += 'NULL'
                            } elseif ($dataType -like '*CHAR*' -or $dataType -like '*CLOB*') {
                                $values += "'" + $value.ToString().Replace("'", "''") + "'"
                            } elseif ($dataType -like '*DATE*' -or $dataType -like '*TIMESTAMP*') {
                                try {
                                    $dt = $reader.GetDateTime($i)
                                    $values += "TO_DATE('" + $dt.ToString('yyyy-MM-dd HH:mm:ss') + "','YYYY-MM-DD HH24:MI:SS')"
                                } catch {
                                    $logLines += "-- 日付変換失敗: $($_.Exception.Message)"
                                    $values += 'NULL'
                                }
                            } else {
                                $values += $value.ToString()
                            }
                        }
                        $insertLines += "INSERT INTO $table (" + ($columnNames -join ', ') + ") VALUES (" + ($values -join ', ') + ");"
                    }

                    if ($FlashbackTimestamp) {
                        $flashbackLines += "-- $counter $table"
                        $flashbackSql = "SELECT " + ($columnNames -join ', ') + " FROM $table AS OF TIMESTAMP TO_TIMESTAMP('$FlashbackTimestamp','YYYY-MM-DD HH24:MI:SS') WHERE $whereClause"
                        if ($orderClause) { $flashbackSql += $orderClause }
                        $flashbackLines += "-- * Captions " + ($columnComments -join ', ')
                        $flashbackLines += $flashbackSql + ';'
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
        @{ Header = 'SELECT'; Lines = $selectLines },
        @{ Header = 'DELETE'; Lines = $deleteLines },
        @{ Header = 'INSERT'; Lines = $insertLines }
    )

    if ($flashbackLines.Count -gt 0) {
        $sections += @{ Header = 'FLASHBACK'; Lines = $flashbackLines }
    }
    if ($logLines.Count -gt 0) {
        $sections += @{ Header = 'LOG / WARNINGS'; Lines = $logLines }
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

        if ($SplitByOperation) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
            if ([string]::IsNullOrWhiteSpace($baseName)) {
                $baseName = 'query'
            }

            $conditionFragment = if ($ConditionSql) { $ConditionSql } else { 'condition' }
            $conditionFragment = Get-SafeFileName -Value $conditionFragment
            if ([string]::IsNullOrWhiteSpace($conditionFragment)) {
                $conditionFragment = 'condition'
            }

            if ($baseName -notlike "*$conditionFragment*") {
                $groupPrefix = "$baseName`_$conditionFragment"
            } else {
                $groupPrefix = $baseName
            }

            foreach ($section in $sections) {
                if (-not $section.Lines -or $section.Lines.Count -eq 0) { continue }

                $sectionLabel = Get-SafeFileName -Value $section.Header.ToLower()
                if ([string]::IsNullOrWhiteSpace($sectionLabel)) {
                    $sectionLabel = 'section'
                }

                $sectionFileName = "$groupPrefix`_$sectionLabel.sql"
                $sectionPath = if ($outputDirectory) {
                    Join-Path $outputDirectory $sectionFileName
                } else {
                    $sectionFileName
                }

                $fileHeaderLine = "-------- $($section.Header) --------"
                $content = New-Object System.Collections.Generic.List[string]
                $content.Add($fileHeaderLine) | Out-Null
                foreach ($line in $section.Lines) {
                    $content.Add($line) | Out-Null
                }

                $content | Set-Content -Path $sectionPath -Encoding UTF8
            }
        } else {
            $outputLines | Set-Content -Path $OutputFile -Encoding UTF8
        }
    }
}
