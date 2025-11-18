param (
    [string]$Column,
    [string]$Condition,
    [string]$ConnectionString = $null,
    [string]$OracleDllPath = $null,
    [string]$OutputDirectory = $null
)

# 任意のPL/SQLプロシージャ呼び出し
# :result OUTパラメータとして受け取る
$plsql = "BEGIN :result := my_package.get_extract_target(); END;"

Write-Host "called."

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptRoot\SelectDeletInsert.ps1"

Write-Host "Column    : $Column"
Write-Host "Condition : $Condition"

if (-not $ConnectionString) {
    $ConnectionString = "User Id=USERID;Password=PASSWORD;Data Source=IPADDRESS:PORT/SID"
}

if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $scriptRoot 'output'
}

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

Write-Host "OutputDir : $OutputDirectory"

$timestamp = Get-Date -Format 'yyyyMMddHHmmss'
$sanitizedCondition = Get-SafeFileName -Value $Condition
if ([string]::IsNullOrWhiteSpace($sanitizedCondition)) {
    $sanitizedCondition = "condition"
}

$baseFileName = "query_{0}_{1}" -f $timestamp, $sanitizedCondition
$outputFile = Join-Path $OutputDirectory ("$baseFileName.sql")

if (-not $OracleDllPath -and $env:ORACLE_DLL_PATH) {
    $OracleDllPath = $env:ORACLE_DLL_PATH
}

$exportParams = @{
    ConnectionString     = $ConnectionString
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
if (Test-Path -Path $OutputDirectory) {
    $filter = "$baseFileName*.sql"
    $outputFiles = Get-ChildItem -Path $OutputDirectory -Filter $filter | Select-Object -ExpandProperty FullName
}

foreach ($file in $outputFiles) {
    Write-Host "Output file : $file"
}

Write-Host "done."
