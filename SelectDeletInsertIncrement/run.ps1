param (
    [string]$Column,
    [string]$Condition,
    [string]$ConnectionString = $null,
    [string]$OracleDllPath = $null,
    [string]$OutputDirectory = $null,
    [string[]]$IncrementColumns = @(),
    [double]$IncrementValue = 1,
    [hashtable]$IncrementOverrides = @{}
)

$plsql = "BEGIN :result := my_package.get_extract_target(); END;"

if (-not $IncrementColumns) {
    $IncrementColumns = @()
}

if (-not $IncrementOverrides) {
    $IncrementOverrides = @{}
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptRoot\SelectDeletInsertIncrement.ps1"

Write-Host "called (increment)."
Write-Host "Column             : $Column"
Write-Host "Condition          : $Condition"
Write-Host "IncrementColumns   : $($IncrementColumns -join ', ')"
Write-Host "IncrementValue     : $IncrementValue"
if ($IncrementOverrides -and $IncrementOverrides.Count -gt 0) {
    $pairs = $IncrementOverrides.GetEnumerator() | ForEach-Object { "${($_.Name)}=${($_.Value)}" }
    Write-Host "IncrementOverrides : $($pairs -join ', ')"
}

if (-not $ConnectionString) {
    $ConnectionString = "User Id=USERID;Password=PASSWORD;Data Source=IPADDRESS:PORT/SID"
}

if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $scriptRoot 'output'
}

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

if (-not $OracleDllPath -and $env:ORACLE_DLL_PATH) {
    $OracleDllPath = $env:ORACLE_DLL_PATH
}

$timestamp = Get-Date -Format 'yyyyMMddHHmmss'
$sanitizedCondition = Get-SafeFileName -Value $Condition
if ([string]::IsNullOrWhiteSpace($sanitizedCondition)) {
    $sanitizedCondition = 'condition'
}

$baseFileName = "query_{0}_{1}" -f $timestamp, $sanitizedCondition
$outputFile = Join-Path $OutputDirectory ("$baseFileName.sql")

$exportParams = @{
    ConnectionString    = $ConnectionString
    Schema              = "SCHEMA"
    TargetColumn        = $Column
    ConditionSql        = $Condition
    FlashbackTimestamp  = "2025-08-17 19:00:00"
    PreSql              = $plsql
    OutputFile          = $outputFile
    SplitByOperation    = $true
    IncrementColumns    = $IncrementColumns
    IncrementValue      = $IncrementValue
    IncrementOverrides  = $IncrementOverrides
}

if ($OracleDllPath) {
    $exportParams.OracleDllPath = $OracleDllPath
}

Export-TableSqlWithIncrementedData @exportParams

if (Test-Path -Path $OutputDirectory) {
    $filter = "$baseFileName*.sql"
    Get-ChildItem -Path $OutputDirectory -Filter $filter | ForEach-Object {
        Write-Host "Output file : $($_.FullName)"
    }
}

Write-Host "done (increment)."
