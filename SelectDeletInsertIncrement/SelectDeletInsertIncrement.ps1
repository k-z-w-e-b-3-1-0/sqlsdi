#requires -Version 3.0

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

function Export-TableSqlWithIncrementedData {
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
        [string]$OutputFile = $null,
        [switch]$SplitByOperation,
        [string[]]$IncrementColumns = @(),
        [double]$IncrementValue = 1,
        [hashtable]$IncrementOverrides = @{}
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

    $sqlDirectory = Resolve-SqlDirectory -StartDirectory $scriptDirectory
    $defaultPrimaryKeySql = Join-Path $sqlDirectory "primary_keys.sql"
    $defaultColumnDefinitionsSql = Join-Path $sqlDirectory "column_definitions.sql"

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

    $invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture

    $incrementMap = @{}
    foreach ($col in $IncrementColumns) {
        if ([string]::IsNullOrWhiteSpace($col)) { continue }
        $incrementMap[$col.ToUpperInvariant()] = $IncrementValue
    }
    foreach ($key in $IncrementOverrides.Keys) {
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $value = $IncrementOverrides[$key]
        try {
            $converted = [System.Convert]::ToDouble($value, $invariantCulture)
        } catch {
            $logLines += "-- Increment override for column '$key' ignored: $($_.Exception.Message)"
            continue
        }
        $incrementMap[$key.ToUpperInvariant()] = $converted
    }

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

        $counter = 1
        foreach ($key in $tables.Keys  | Sort-Object) {
            $meta = $tables[$key]
            $colNames = $meta | ForEach-Object { $_[0] }
            $flatCols = $colNames
            $flatComments = $meta | ForEach-Object { $_[1] }
            $flatTypes = $meta | ForEach-Object { $_[2] }
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
                            $columnName = $flatCols[$i]
                            $typeInfo = $flatTypes[$i]
                            $v = $reader.GetValue($i)

                            $columnKey = $columnName.ToUpperInvariant()
                            $shouldIncrement = $incrementMap.ContainsKey($columnKey)
                            $incrementAmount = if ($shouldIncrement) { $incrementMap[$columnKey] } else { $null }

                            if ($v -eq $null -or $reader.IsDBNull($i)) {
                                if ($shouldIncrement) {
                                    $logLines += "-- Increment skipped for $key.$columnName: value is NULL"
                                }
                                $vals += "NULL"
                                continue
                            }

                            if ($shouldIncrement) {
                                if ($typeInfo -notmatch 'NUMBER|FLOAT|DECIMAL|INT|BINARY_DOUBLE|BINARY_FLOAT') {
                                    $logLines += "-- Increment skipped for $key.$columnName: data type '$typeInfo' is not numeric"
                                    $incrementAmount = $null
                                }
                            }

                            if ($shouldIncrement -and $incrementAmount -ne $null) {
                                try {
                                    $baseValue = [System.Convert]::ToDouble($v, $invariantCulture)
                                    $newValue = $baseValue + [System.Convert]::ToDouble($incrementAmount, $invariantCulture)
                                    $vals += $newValue.ToString('G', $invariantCulture)
                                    continue
                                } catch {
                                    $logLines += "-- Increment failed for $key.$columnName: $($_.Exception.Message)"
                                }
                            }

                            if ($typeInfo -like "*CHAR*" -or $typeInfo -like "*CLOB*") {
                                $vals += "'" + $v.ToString().Replace("'", "''") + "'"
                            } elseif ($typeInfo -like "*DATE*" -or $typeInfo -like "*TIMESTAMP*") {
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
        if ($SplitByOperation) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputFile)
            if ([string]::IsNullOrWhiteSpace($baseName)) {
                $baseName = "query"
            }

            $conditionFragment = Get-SafeFileName -Value $ConditionSql
            if ([string]::IsNullOrWhiteSpace($conditionFragment)) {
                $conditionFragment = "condition"
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
                    $sectionLabel = "section"
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
