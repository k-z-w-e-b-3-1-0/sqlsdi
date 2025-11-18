cd %~dp0

@echo off

set NOW=%date:~0,4%%date:~5,2%%date:~8,2%%time:~0,2%%time:~3,2%%time:~6,2%
echo %NOW%

REM Optionally set ORACLE_DLL_PATH before invoking PowerShell, for example:
REM set ORACLE_DLL_PATH=C:\path\to\Oracle.DataAccess.dll

C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SelectDeletInsert\run.ps1 "column_name" "condition"

pause
