function Export-TableSqlWithData {
    param(
        [string]$ConnectionString,
        [string]$Schema,
        [string]$TargetColumn,
        [string]$ConditionSql,
        [string]$FlashbackTimestamp = $null,
        [string]$PreSql = $null
    )

    [Reflection.Assembly]::LoadFile("C:\somewhere\Oracle.DataAccess.dll")
    $conn = New-Object Oracle.DataAccess.Client.OracleConnection($ConnectionString)
    $conn.Open()

    $selectLines     = @()
    $deleteLines     = @()
    $insertLines     = @()
    $flashbackLines  = @()
    $logLines        = @()

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
    $pkSql = @"
SELECT acc.OWNER, acc.TABLE_NAME, acc.COLUMN_NAME, acc.POSITION
FROM USER_CONS_COLUMNS acc
JOIN USER_CONSTRAINTS ac ON acc.OWNER = ac.OWNER AND acc.CONSTRAINT_NAME = ac.CONSTRAINT_NAME
WHERE ac.CONSTRAINT_TYPE = 'P'
AND acc.OWNER = '$Schema'
ORDER BY acc.OWNER, acc.TABLE_NAME, acc.POSITION
"@
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $pkSql
    $reader = $cmd.ExecuteReader()
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
    $reader.Close()

    # 列定義取得
    $defSql = @"
SELECT '$Schema', c.TABLE_NAME, c.COLUMN_NAME, com.COMMENTS, c.COLUMN_ID, c.DATA_TYPE
FROM USER_TAB_COLUMNS c
LEFT JOIN USER_COL_COMMENTS com
  ON c.TABLE_NAME = com.TABLE_NAME AND c.COLUMN_NAME = com.COLUMN_NAME
WHERE EXISTS (
  SELECT 1 FROM USER_TAB_COLUMNS t
  WHERE t.TABLE_NAME = c.TABLE_NAME AND t.COLUMN_NAME = '$TargetColumn'
)
AND EXISTS (
   SELECT 1 FROM USER_TABLES u
   WHERE u.TABLE_NAME = c.TABLE_NAME
)
ORDER BY c.TABLE_NAME, c.COLUMN_ID
"@
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $defSql
    $tables = @{}
    $reader = $cmd.ExecuteReader()
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
    $reader.Close()

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

#        $selectSql = "SELECT " + ($flatCols -join ", ") + " FROM $key WHERE $whereClause" + $orderClause + ";"
#        $selectLines += "-- SELECT for $key"
#        $selectLines += $selectSql

#        $deleteLines += "-- DELETE for $key"
#        $deleteLines += "DELETE FROM $key WHERE $whereClause;"

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT " + ($flatCols -join ", ") + " FROM $key WHERE $whereClause"
        
        $reader = $cmd.ExecuteReader()

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
        $reader.Close()
    }

    $conn.Close()

    # 出力
    Write-Output "`n-------- SELECT --------"
    $selectLines | ForEach-Object { Write-Output $_ }
    Write-Output "`n-------- DELETE --------"
    $deleteLines | ForEach-Object { Write-Output $_ }
    Write-Output "`n-------- INSERT --------"
    $insertLines | ForEach-Object { Write-Output $_ }
    if ($flashbackLines.Count -gt 0) {
        Write-Output "`n-------- FLASHBACK --------"
        $flashbackLines | ForEach-Object { Write-Output $_ }
    }
    if ($logLines.Count -gt 0) {
        Write-Output "`n-------- LOG / WARNINGS --------"
        $logLines | ForEach-Object { Write-Output $_ }
    }
}
