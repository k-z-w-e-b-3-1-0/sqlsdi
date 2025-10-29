param (
    [string]$Column,
    [string]$Condition
)

# 任意のPL/SQプロシージャ呼び出し
# :result OUTパラメータとして受け取る
$plsql = "BEGIN :result := my_package.get_extract_target(); END;"

    Write-Host "called."

. .\SelectDeleteInsert.ps1

    Write-Host "Column    : $Column"
    Write-Host "Condition : $Condition"

Export-TableSqlWithData -ConnectionString "User Id=USERID;Password=PASSWORD;Data Source=IPADDRESS:PORT/SID" `
                         -Schema "SCHEMA" `
                         -TargetColumn $Column `
                         -ConditionSql $Condition `
                         -FlashbackTimestamp "2025-08-17 19:00:00" `
                         -PreSql $plsql

    Write-Host "done."

