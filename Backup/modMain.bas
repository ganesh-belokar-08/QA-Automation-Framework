Attribute VB_Name = "modMain"
Option Explicit

'=========================================================
' QA Automation Framework
' Main Entry Point
'=========================================================

Public Sub StartFramework()

    Dim TargetWB As Workbook
    Dim FileName As String
    Dim TxtFileName As String

    On Error GoTo ErrorHandler

    '----------------------------------------
    ' Improve Performance
    '----------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Starting QA Automation..."

    '----------------------------------------
    ' Initialize Framework
    '----------------------------------------
    InitializePaths

    ValidateFolders

    '----------------------------------------
    ' Get Workbook
    '----------------------------------------
    FileName = GetInputWorkbook

    If FileName = "" Then GoTo ExitHandler

    '----------------------------------------
    ' Open Workbook
    '----------------------------------------
    Set TargetWB = Workbooks.Open(InputPath & FileName)

    '----------------------------------------
    ' Report Name
    '----------------------------------------
    InputWorkbookName = Left(FileName, InStrRev(FileName, ".") - 1)

    TimeStamp = Format(Now, "yyyymmdd_hhnnss")

    ReportFileName = Replace(InputWorkbookName, "-MATRIX", "") & _
                     "-Report_" & TimeStamp & ".xlsx"

    '----------------------------------------
    ' Run Framework
    '----------------------------------------
    ProcessingStart = Now

    RunFramework TargetWB

    ProcessingEnd = Now

    '----------------------------------------
    ' Save Workbook
    '----------------------------------------
    TargetWB.Save

    '----------------------------------------
    'Create Report Copy
    '----------------------------------------
    TargetWB.SaveCopyAs ReportPath & ReportFileName

    '----------------------------------------
    ' Create .txt Summary (same name as report)
    '----------------------------------------
    TxtFileName = Left(ReportFileName, InStrRev(ReportFileName, ".") - 1) & ".txt"

    WriteSummaryTextFile ReportPath & TxtFileName

    '----------------------------------------
    'Completed
    '----------------------------------------
    MsgBox _
    "QA Automation Completed Successfully." & vbCrLf & vbCrLf & _
    "Workbook : " & InputWorkbookName & vbCrLf & _
    "Report : " & ReportFileName & vbCrLf & _
    "Total TC : " & TotalTC & vbCrLf & _
    "Passed : " & PassedTC & vbCrLf & _
    "Failed : " & FailedTC & vbCrLf & _
    IIf(FailedTC > 0, "Failed TC(s) : " & FailedTCList & vbCrLf, "") & _
    "Time Taken : " & _
    Format(ProcessingEnd - ProcessingStart, "hh:mm:ss"), _
    vbInformation

ExitHandler:

    On Error Resume Next

    If Not TargetWB Is Nothing Then
        TargetWB.Close SaveChanges:=True
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Exit Sub

ErrorHandler:

    MsgBox _
    "Error " & Err.Number & vbCrLf & _
    Err.Description, vbCritical

    Resume ExitHandler

End Sub

'=========================================================
' Initialize Framework Paths
'=========================================================

Private Sub InitializePaths()

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the Framework workbook first.", vbCritical
        End
    End If

    'Current folder
    FrameworkPath = ThisWorkbook.Path

    'Remove \Source
    If LCase$(Right$(FrameworkPath, 7)) = "\source" Then
        FrameworkPath = Left$(FrameworkPath, Len(FrameworkPath) - 7)
    End If

    InputPath = FrameworkPath & "\Input\"
    OutputPath = FrameworkPath & "\Output\"
    SplitPath = FrameworkPath & "\Split\"
    BackupPath = FrameworkPath & "\Backup\"
    ReportPath = FrameworkPath & "\Reports\"

End Sub

'=========================================================
' Returns Workbook Name
'=========================================================

Private Function GetInputWorkbook() As String

    Dim FileName As String
    Dim Count As Long

    FileName = Dir(InputPath & "*.xls*")

    Do While FileName <> ""

        Count = Count + 1

        If Count = 1 Then
            GetInputWorkbook = FileName
        End If

        FileName = Dir

    Loop

    Select Case Count

        Case 0

            MsgBox "No workbook found in Input folder.", vbExclamation
            GetInputWorkbook = ""

        Case 1

            'Valid

        Case Else

            MsgBox _
            "Multiple workbooks found in Input folder." & vbCrLf & _
            "Please keep only ONE workbook.", vbCritical

            GetInputWorkbook = ""

    End Select

End Function

