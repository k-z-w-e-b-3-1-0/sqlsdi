param (
    [Parameter(Mandatory = $true)]
    [string[]]$Columns,

    [string]$Condition,

    [string[]]$ConditionFragments,

    [hashtable]$EqualConditions,

    [string]$OracleDllPath = $null
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptRoot\SelectDeletInsert.ps1"
. "$scriptRoot\SelectDeletInsertMulti.ps1"

Write-Host "called (multi-column)."
Write-Host "Columns             : $($Columns -join ', ')"
Write-Host "Condition           : $Condition"
if ($ConditionFragments) {
    Write-Host "ConditionFragments : $($ConditionFragments -join ' | ')"
}
if ($EqualConditions) {
    $pairs = $EqualConditions.GetEnumerator() | ForEach-Object { "${($_.Name)}=${($_.Value)}" }
    Write-Host "EqualConditions    : $($pairs -join ', ')"
}

$timestamp = Get-Date -Format 'yyyyMMddHHmmss'

$conditionLabel = if (-not [string]::IsNullOrWhiteSpace($Condition)) {
    $Condition
} elseif ($ConditionFragments -and $ConditionFragments.Count -gt 0) {
    $ConditionFragments -join ' AND '
} elseif ($EqualConditions -and $EqualConditions.Count -gt 0) {
    ($EqualConditions.GetEnumerator() | ForEach-Object { "${($_.Name)}=${($_.Value)}" }) -join ' AND '
} else {
    'condition'
}

$sanitizedCondition = Get-SafeFileName -Value $conditionLabel
if ([string]::IsNullOrWhiteSpace($sanitizedCondition)) {
    $sanitizedCondition = 'condition'
}

$baseFileName = "query_{0}_{1}" -f $timestamp, $sanitizedCondition
$outputDirectory = Join-Path $scriptRoot 'output'
$outputFile = Join-Path $outputDirectory ("$baseFileName.sql")

if (-not $OracleDllPath -and $env:ORACLE_DLL_PATH) {
    $OracleDllPath = $env:ORACLE_DLL_PATH
}

$exportParams = @{
    ConnectionString    = "User Id=USERID;Password=PASSWORD;Data Source=IPADDRESS:PORT/SID"
    Schema              = "SCHEMA"
    TargetColumns       = $Columns
    ConditionSql        = $Condition
    ConditionFragments  = $ConditionFragments
    EqualityConditions  = $EqualConditions
    FlashbackTimestamp  = "2025-08-17 19:00:00"
    PreSql              = "BEGIN :result := my_package.get_extract_target(); END;"
    OutputFile          = $outputFile
    SplitByOperation    = $true
}

if ($OracleDllPath) {
    $exportParams.OracleDllPath = $OracleDllPath
}

Export-TableSqlWithDataMultiColumn @exportParams

if (Test-Path -Path $outputDirectory) {
    $filter = "$baseFileName*.sql"
    Get-ChildItem -Path $outputDirectory -Filter $filter | ForEach-Object {
        Write-Host "Output file : $($_.FullName)"
    }
}

Write-Host "done (multi-column)."
