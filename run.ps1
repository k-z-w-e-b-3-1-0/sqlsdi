param (
    [string]$Column,
    [string]$Condition,
    [string]$OracleDllPath = $null
)

# 任意のPL/SQLプロシージャ呼び出し
# :result OUTパラメータとして受け取る
$plsql = "BEGIN :result := my_package.get_extract_target(); END;"

Write-Host "called."

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptRoot\SelectDeletInsert.ps1"

Write-Host "Column    : $Column"
Write-Host "Condition : $Condition"

$timestamp = Get-Date -Format 'yyyyMMddHHmmss'
$sanitizedCondition = Get-SafeFileName -Value $Condition
if ([string]::IsNullOrWhiteSpace($sanitizedCondition)) {
    $sanitizedCondition = "condition"
}

$baseFileName = "query_{0}_{1}" -f $timestamp, $sanitizedCondition
$outputDirectory = Join-Path $scriptRoot 'output'
$outputFile = Join-Path $outputDirectory ("$baseFileName.sql")

if (-not $OracleDllPath -and $env:ORACLE_DLL_PATH) {
    $OracleDllPath = $env:ORACLE_DLL_PATH
}

$exportParams = @{
    ConnectionString     = "User Id=USERID;Password=PASSWORD;Data Source=IPADDRESS:PORT/SID"
    Schema               = "SCHEMA"
    TargetColumn         = $Column
    ConditionSql         = $Condition
    FlashbackTimestamp   = "2025-08-17 19:00:00"
    PreSql               = $plsql
    OutputFile           = $outputFile
    SplitByOperation     = $true
}

if ($OracleDllPath) {
    $exportParams.OracleDllPath = $OracleDllPath
}

Export-TableSqlWithData @exportParams

$outputFiles = @()
if (Test-Path -Path $outputDirectory) {
    $filter = "$baseFileName*.sql"
    $outputFiles = Get-ChildItem -Path $outputDirectory -Filter $filter | Select-Object -ExpandProperty FullName
}

foreach ($file in $outputFiles) {
    Write-Host "Output file : $file"
}

Write-Host "done."
