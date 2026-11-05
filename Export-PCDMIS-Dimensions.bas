' Copyright (C) 2026 Townsend Precision Labs Ltd
' SPDX-License-Identifier: GPL-3.0-or-later
' PC-DMIS in-process fallback dimension exporter.
' Insert this file as a BASIC Script command after dimensions execute and before LOOP/END.

Function CleanField(Value)
    Dim TextValue As String
    Dim CleanValue As String
    Dim CharacterValue As String
    Dim CharacterCode As Integer
    Dim CharacterIndex As Long

    TextValue = CStr(Value)
    CleanValue = ""
    For CharacterIndex = 1 To Len(TextValue)
        CharacterValue = Mid(TextValue, CharacterIndex, 1)
        CharacterCode = Asc(CharacterValue)
        If CharacterCode = 9 Or CharacterCode = 10 Or CharacterCode = 13 Then
            CharacterValue = " "
        End If
        CleanValue = CleanValue & CharacterValue
    Next CharacterIndex
    CleanField = CleanValue
End Function

Sub ExportDimensions(OutputPath As String, OperatorName As String, JobNumber As String, SerialNumber As String, PartNumber As String, Revision As String)
    Dim App As Object
    Dim PartProgram As Object
    Dim Commands As Object
    Dim Command As Object
    Dim ExecutionCommand As Object
    Dim Dimension As Object
    Dim CommentCommand As Object
    Dim ShellObject As Object
    Dim ApplicationId As String
    Dim CommandCount As Long
    Dim CommandIndex As Long
    Dim ExecutionCount As Long
    Dim ItemNumber As Long
    Dim BridgeFile As Integer
    Dim EditWindowTextFile As Integer
    Dim RequestFile As Integer
    Dim ErrorFile As Integer
    Dim ReturnCode As Long
    Dim CurrentDimensionId As String
    Dim CurrentReportLabel As String
    Dim DimensionId As String
    Dim OwnId As String
    Dim AxisText As String
    Dim Features As String
    Dim StatusText As String
    Dim HeaderText As String
    Dim RowText As String
    Dim EditWindowText As String
    Dim BridgePath As String
    Dim EditWindowTextPath As String
    Dim RequestPath As String
    Dim LauncherPath As String
    Dim Stage As String
    Dim OutTolerance As Double

    On Error GoTo Failed

    BridgePath = "C:\PCDMIS_Report_Fallback\last-basic-bridge.tsv"
    EditWindowTextPath = "C:\PCDMIS_Report_Fallback\last-command-mode.txt"
    RequestPath = "C:\PCDMIS_Report_Fallback\export-request.txt"
    LauncherPath = "C:\PCDMIS_Report_Fallback\Export-PCDMIS-Dimensions.ps1"

    Stage = "read intrinsic versioned PC-DMIS application ID"
    ApplicationId = CStr(Versioned_PCDLRN_Application_ID)
    Stage = "create intrinsic versioned PC-DMIS application object"
    Set App = CreateObject(ApplicationId)
    Stage = "validate PC-DMIS object"
    If App Is Nothing Then
        MsgBox "The BASIC exporter could not attach to the running PC-DMIS application.", 16, "PC-DMIS report fallback"
        Exit Sub
    End If

    Stage = "read active part program"
    Set PartProgram = App.ActivePartProgram
    If PartProgram Is Nothing Then
        MsgBox "PC-DMIS did not expose the active measurement routine to the BASIC exporter.", 16, "PC-DMIS report fallback"
        Exit Sub
    End If

    Stage = "read command collection"
    Set Commands = PartProgram.Commands
    CommandCount = Commands.Count

    Stage = "open bridge file"
    BridgeFile = FreeFile
    Open BridgePath For Output As #BridgeFile
    HeaderText = "Item" & Chr(9) & "ReportLabel" & Chr(9) & "Dimension"
    HeaderText = HeaderText & Chr(9) & "Axis" & Chr(9) & "Features" & Chr(9) & "Nominal"
    HeaderText = HeaderText & Chr(9) & "PlusTolerance" & Chr(9) & "MinusTolerance"
    HeaderText = HeaderText & Chr(9) & "Measured" & Chr(9) & "Deviation" & Chr(9) & "OutTolerance"
    HeaderText = HeaderText & Chr(9) & "Bonus" & Chr(9) & "Minimum" & Chr(9) & "Maximum"
    HeaderText = HeaderText & Chr(9) & "Status" & Chr(9) & "Execution"
    HeaderText = HeaderText & Chr(9) & "CommandIndex" & Chr(9) & "CommandType"
    Print #BridgeFile, HeaderText

    CurrentDimensionId = ""
    CurrentReportLabel = ""
    ItemNumber = 0

    For CommandIndex = 1 To CommandCount
        Stage = "read command " & CStr(CommandIndex)
        On Error GoTo SkipCommand
        Set Command = Commands.Item(CommandIndex)

        If Command.IsComment Then
            Set CommentCommand = Command.CommentCommand
            If CommentCommand.CommentType = 1 Then
                CurrentReportLabel = CleanField(CommentCommand.Comment)
            End If
            GoTo NextCommand
        End If

        If Command.IsDimension = False Then GoTo NextCommand

        ExecutionCount = Command.Count
        If ExecutionCount < 1 Then ExecutionCount = 1
        Set ExecutionCommand = Command.Item(ExecutionCount)
        Set Dimension = ExecutionCommand.DimensionCommand

        OwnId = CleanField(Dimension.ID)
        AxisText = CleanField(Dimension.AxisLetter)
        If OwnId <> "" Then CurrentDimensionId = OwnId
        If AxisText = "" Then GoTo NextCommand

        DimensionId = OwnId
        If DimensionId = "" Then DimensionId = CurrentDimensionId
        If DimensionId = "" Then DimensionId = "DIMENSION_AT_COMMAND_" & CStr(CommandIndex)

        Features = CleanField(Dimension.Feat1)
        If CleanField(Dimension.Feat2) <> "" Then
            If Features <> "" Then Features = Features & "; "
            Features = Features & CleanField(Dimension.Feat2)
        End If
        If CleanField(Dimension.Feat3) <> "" Then
            If Features <> "" Then Features = Features & "; "
            Features = Features & CleanField(Dimension.Feat3)
        End If

        OutTolerance = CDbl(Dimension.OutTol)
        StatusText = "PASS"
        If Abs(OutTolerance) > 0.0000001 Then StatusText = "FAIL"

        ItemNumber = ItemNumber + 1
        RowText = CStr(ItemNumber)
        RowText = RowText & Chr(9) & CleanField(CurrentReportLabel)
        RowText = RowText & Chr(9) & CleanField(DimensionId)
        RowText = RowText & Chr(9) & AxisText
        RowText = RowText & Chr(9) & Features
        RowText = RowText & Chr(9) & CStr(Dimension.Nominal)
        RowText = RowText & Chr(9) & CStr(Abs(Dimension.Plus))
        RowText = RowText & Chr(9) & CStr(Abs(Dimension.Minus))
        RowText = RowText & Chr(9) & CStr(Dimension.Measured)
        RowText = RowText & Chr(9) & CStr(Dimension.Deviation)
        RowText = RowText & Chr(9) & CStr(OutTolerance)
        RowText = RowText & Chr(9) & CStr(Dimension.Bonus)
        RowText = RowText & Chr(9) & CStr(Dimension.Min)
        RowText = RowText & Chr(9) & CStr(Dimension.Max)
        RowText = RowText & Chr(9) & StatusText
        RowText = RowText & Chr(9) & CStr(ExecutionCount)
        RowText = RowText & Chr(9) & CStr(CommandIndex)
        RowText = RowText & Chr(9) & "BASIC_" & CStr(ExecutionCommand.Type)
        Print #BridgeFile, RowText

NextCommand:
        On Error GoTo Failed
    Next CommandIndex

    Close #BridgeFile
    If ItemNumber = 0 Then
        MsgBox "The BASIC exporter found no evaluated legacy dimension rows.", 16, "PC-DMIS report fallback"
        Exit Sub
    End If

    Stage = "read command-mode Edit window text"
    EditWindowText = CStr(PartProgram.EditWindowTextAll)
    Stage = "write command-mode Edit window text"
    EditWindowTextFile = FreeFile
    Open EditWindowTextPath For Output As #EditWindowTextFile
    Print #EditWindowTextFile, EditWindowText
    Close #EditWindowTextFile

    RequestFile = FreeFile
    Open RequestPath For Output As #RequestFile
    Print #RequestFile, OutputPath
    Print #RequestFile, OperatorName
    Print #RequestFile, JobNumber
    Print #RequestFile, SerialNumber
    Print #RequestFile, PartNumber
    Print #RequestFile, Revision
    Print #RequestFile, EditWindowTextPath
    Print #RequestFile, BridgePath
    Close #RequestFile

    Set ShellObject = CreateObject("WScript.Shell")
    ReturnCode = ShellObject.Run("powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & LauncherPath & Chr(34), 0, True)
    If ReturnCode <> 0 Then
        MsgBox "The Excel fallback exporter failed. Check C:\PCDMIS_Report_Fallback\last-export-error.log", 16, "PC-DMIS report fallback"
    End If
    Exit Sub

SkipCommand:
    Err.Clear
    Resume NextCommand

Failed:
    On Error Resume Next
    Close #BridgeFile
    Close #EditWindowTextFile
    Close #RequestFile
    ErrorFile = FreeFile
    Open "C:\PCDMIS_Report_Fallback\last-basic-error.log" For Output As #ErrorFile
    Print #ErrorFile, "Stage: " & Stage
    Print #ErrorFile, "Error " & CStr(Err.Number) & ": " & Err.Description
    Close #ErrorFile
    MsgBox "The in-process PC-DMIS exporter failed. Check C:\PCDMIS_Report_Fallback\last-basic-error.log", 16, "PC-DMIS report fallback"
End Sub
