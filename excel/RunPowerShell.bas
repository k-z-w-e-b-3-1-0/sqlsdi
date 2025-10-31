Attribute VB_Name = "PowerShellBatch"
Option Explicit

Private Const POWERSHELL_EXE As String = "powershell.exe"
Private Const PARAMETER_SHEET_NAME As String = "Sheet1"
Private Const FIRST_DATA_ROW As Long = 2

Private Const COL_TARGET_COLUMN As Long = 1 ' A 列
Private Const COL_CONDITION As Long = 2    ' B 列
Private Const COL_CONNECTION_STRING As Long = 3 ' C 列
Private Const COL_DLL_PATH As Long = 4          ' D 列
Private Const COL_OUTPUT_DIRECTORY As Long = 5  ' E 列

Private Const DEFAULT_SCRIPT_NAME As String = "run.ps1"
Private Const SCRIPT_PATH_NAME As String = "PowerShellScriptPath"
Private Const DEFAULT_CONNECTION_NAME As String = "DefaultConnectionString"
Private Const DEFAULT_DLL_NAME As String = "DefaultOracleDllPath"
Private Const DEFAULT_OUTPUT_NAME As String = "DefaultOutputDirectory"

' Excel の一覧から run.ps1 を同期的に実行します。
Public Sub RunPowerShellBatch()
    RunPowerShellBatchInternal True
End Sub

' Excel の一覧から run.ps1 を非同期に実行します（実行完了を待ちません）。
Public Sub RunPowerShellBatchAsync()
    RunPowerShellBatchInternal False
End Sub

' 現在選択中の行だけを実行します。
Public Sub RunPowerShellForActiveRow()
    Dim activeRow As Long
    activeRow = ActiveCell.Row

    If activeRow < FIRST_DATA_ROW Then
        MsgBox "ヘッダーではなく、対象データのセルを選択してください。", vbInformation
        Exit Sub
    End If

    RunSingleRow activeRow, True
End Sub

Private Sub RunPowerShellBatchInternal(ByVal waitForCompletion As Boolean)
    Dim sheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim shell As Object

    Set sheet = GetParameterSheet()
    If sheet Is Nothing Then
        MsgBox "パラメータ入力用のシート (" & PARAMETER_SHEET_NAME & ") が見つかりません。", vbCritical
        Exit Sub
    End If

    lastRow = sheet.Cells(sheet.Rows.Count, COL_TARGET_COLUMN).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then
        MsgBox "処理対象の行がありません。", vbInformation
        Exit Sub
    End If

    On Error GoTo ShellError
    Set shell = CreateObject("WScript.Shell")

    For rowIndex = FIRST_DATA_ROW To lastRow
        RunSingleRow rowIndex, waitForCompletion, shell
    Next rowIndex
    Exit Sub

ShellError:
    MsgBox "PowerShell の起動に失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub RunSingleRow(ByVal rowIndex As Long, ByVal waitForCompletion As Boolean, Optional ByVal existingShell As Object)
    Dim sheet As Worksheet
    Dim columnValue As String
    Dim conditionValue As String
    Dim connectionValue As String
    Dim dllValue As String
    Dim outputValue As String
    Dim command As String
    Dim shell As Object

    Set sheet = GetParameterSheet()
    If sheet Is Nothing Then Exit Sub

    columnValue = Trim$(CStr(sheet.Cells(rowIndex, COL_TARGET_COLUMN).Value))
    conditionValue = Trim$(CStr(sheet.Cells(rowIndex, COL_CONDITION).Value))
    connectionValue = ResolveRowValue(sheet.Cells(rowIndex, COL_CONNECTION_STRING).Value, DEFAULT_CONNECTION_NAME)
    dllValue = ResolvePathValue(sheet.Cells(rowIndex, COL_DLL_PATH).Value, DEFAULT_DLL_NAME)
    outputValue = ResolvePathValue(sheet.Cells(rowIndex, COL_OUTPUT_DIRECTORY).Value, DEFAULT_OUTPUT_NAME)

    If Len(columnValue) = 0 And Len(conditionValue) = 0 Then
        Exit Sub
    End If

    command = BuildCommand(columnValue, conditionValue, connectionValue, dllValue, outputValue)
    If Len(command) = 0 Then
        Exit Sub
    End If

    On Error GoTo ShellError

    If existingShell Is Nothing Then
        Set shell = CreateObject("WScript.Shell")
    Else
        Set shell = existingShell
    End If

    shell.Run command, 0, waitForCompletion
    Exit Sub

ShellError:
    MsgBox "PowerShell の起動に失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub

Private Function BuildCommand(ByVal columnValue As String, _
                              ByVal conditionValue As String, _
                              ByVal connectionValue As String, _
                              ByVal dllValue As String, _
                              ByVal outputValue As String) As String
    Dim scriptPath As String
    Dim command As String

    scriptPath = ResolveScriptPath()
    If Len(scriptPath) = 0 Then
        MsgBox "run.ps1 のパスが設定されていません。名前付きセル PowerShellScriptPath を追加するか、ブックと同じフォルダーに run.ps1 を配置してください。", vbCritical
        Exit Function
    End If

    If Len(columnValue) = 0 Then
        MsgBox "列名が空の行はスキップされました。", vbExclamation
        Exit Function
    End If

    If Len(conditionValue) = 0 Then
        MsgBox "条件 SQL が空の行はスキップされました。", vbExclamation
        Exit Function
    End If

    command = POWERSHELL_EXE & " -NoProfile -ExecutionPolicy Bypass" _
              & " -File " & QuoteForCommand(scriptPath) _
              & " -Column " & QuoteForCommand(columnValue) _
              & " -Condition " & QuoteForCommand(conditionValue)

    If Len(connectionValue) > 0 Then
        command = command & " -ConnectionString " & QuoteForCommand(connectionValue)
    End If

    If Len(dllValue) > 0 Then
        command = command & " -OracleDllPath " & QuoteForCommand(dllValue)
    End If

    If Len(outputValue) > 0 Then
        command = command & " -OutputDirectory " & QuoteForCommand(outputValue)
    End If

    BuildCommand = command
End Function

Private Function GetParameterSheet() As Worksheet
    On Error Resume Next
    Set GetParameterSheet = ThisWorkbook.Worksheets(PARAMETER_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function ResolveScriptPath() As String
    Dim specifiedPath As String

    specifiedPath = ResolveRowValue(vbNullString, SCRIPT_PATH_NAME)
    If Len(specifiedPath) = 0 Then
        If Len(ThisWorkbook.Path) > 0 Then
            specifiedPath = CombinePath(ThisWorkbook.Path, DEFAULT_SCRIPT_NAME)
        Else
            specifiedPath = DEFAULT_SCRIPT_NAME
        End If
    End If

    specifiedPath = NormalizePath(specifiedPath)

    ResolveScriptPath = specifiedPath
End Function

Private Function ResolveRowValue(ByVal cellValue As Variant, ByVal fallbackName As String) As String
    Dim trimmedValue As String

    trimmedValue = Trim$(CStr(cellValue))
    If Len(trimmedValue) > 0 Then
        ResolveRowValue = trimmedValue
        Exit Function
    End If

    ResolveRowValue = ResolveNamedValue(fallbackName)
End Function

Private Function ResolvePathValue(ByVal cellValue As Variant, ByVal fallbackName As String) As String
    Dim pathValue As String

    pathValue = ResolveRowValue(cellValue, fallbackName)
    If Len(pathValue) > 0 Then
        ResolvePathValue = NormalizePath(pathValue)
    Else
        ResolvePathValue = ""
    End If
End Function

Private Function ResolveNamedValue(ByVal name As String) As String
    Dim targetRange As Range

    On Error Resume Next
    Set targetRange = ThisWorkbook.Names(name).RefersToRange
    On Error GoTo 0

    If Not targetRange Is Nothing Then
        ResolveNamedValue = Trim$(CStr(targetRange.Value))
    Else
        ResolveNamedValue = ""
    End If
End Function

Private Function NormalizePath(ByVal value As String) As String
    Dim expanded As String

    expanded = ExpandEnvironmentStrings(value)
    If IsAbsolutePath(expanded) Then
        NormalizePath = expanded
    ElseIf Len(ThisWorkbook.Path) > 0 Then
        NormalizePath = CombinePath(ThisWorkbook.Path, expanded)
    Else
        NormalizePath = expanded
    End If
End Function

Private Function ExpandEnvironmentStrings(ByVal value As String) As String
    Dim wsh As Object

    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        ExpandEnvironmentStrings = value
        Err.Clear
    Else
        ExpandEnvironmentStrings = wsh.ExpandEnvironmentStrings(value)
    End If
    On Error GoTo 0
End Function

Private Function CombinePath(ByVal basePath As String, ByVal relativePath As String) As String
    If Right$(basePath, 1) = Application.PathSeparator Then
        CombinePath = basePath & relativePath
    Else
        CombinePath = basePath & Application.PathSeparator & relativePath
    End If
End Function

Private Function IsAbsolutePath(ByVal value As String) As Boolean
    If Len(value) = 0 Then
        IsAbsolutePath = False
    ElseIf Left$(value, 2) = "\\" Then
        IsAbsolutePath = True
    ElseIf Mid$(value, 2, 1) = ":" Then
        IsAbsolutePath = True
    ElseIf Left$(value, 1) = "/" Then
        IsAbsolutePath = True
    Else
        IsAbsolutePath = False
    End If
End Function

Private Function QuoteForCommand(ByVal value As String) As String
    QuoteForCommand = "\"" & Replace(value, "\"", "\"\"") & "\""
End Function
