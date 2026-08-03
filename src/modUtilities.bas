Attribute VB_Name = "modUtilities"
Option Explicit

'=========================================================
' Check Whether Sheet is a Test Case Sheet
' Only sheets named TC1, TC01, TC10, etc. are treated as TC sheets
'=========================================================

Public Function IsTCSheet(ws As Worksheet) As Boolean

    If UCase(ws.Name) = UCase(MATRIX_SHEET) Then
        IsTCSheet = False
        Exit Function
    End If

    If UCase(ws.Name) = UCase(REPORT_SHEET) Then
        IsTCSheet = False
        Exit Function
    End If

    If Left(UCase(ws.Name), 2) <> "TC" Then
        IsTCSheet = False
        Exit Function
    End If

    IsTCSheet = True

End Function

'=========================================================
' Check Whether Folder Exists
'=========================================================

Public Function FolderExists(ByVal FolderPath As String) As Boolean

    On Error Resume Next
    FolderExists = (Len(Dir(FolderPath, vbDirectory)) > 0)
    On Error GoTo 0

End Function

'=========================================================
' Create Folder
'=========================================================

Public Sub CreateFolder(ByVal FolderPath As String)

    If Not FolderExists(FolderPath) Then
        MkDir FolderPath
    End If

End Sub

'=========================================================
' Validate Framework Folder Structure
'=========================================================

Public Sub ValidateFolders()

    Debug.Print FrameworkPath
    Debug.Print InputPath
    Debug.Print OutputPath
    Debug.Print SplitPath
    Debug.Print BackupPath
    Debug.Print ReportPath

    If Dir(InputPath, vbDirectory) = "" Then MkDir InputPath
    If Dir(OutputPath, vbDirectory) = "" Then MkDir OutputPath
    If Dir(SplitPath, vbDirectory) = "" Then MkDir SplitPath
    If Dir(BackupPath, vbDirectory) = "" Then MkDir BackupPath
    If Dir(ReportPath, vbDirectory) = "" Then MkDir ReportPath

End Sub

'=========================================================
' Write Execution Summary as .txt File
'=========================================================

Public Sub WriteSummaryTextFile(ByVal FilePath As String)

    Dim FileNum As Integer
    FileNum = FreeFile

    Open FilePath For Output As #FileNum

    Print #FileNum, "====================================================="
    Print #FileNum, " QA AUTOMATION FRAMEWORK - EXECUTION SUMMARY"
    Print #FileNum, "====================================================="
    Print #FileNum, "Workbook       : " & InputWorkbookName
    Print #FileNum, "Report File    : " & ReportFileName
    Print #FileNum, "Date & Time    : " & Format(Now, "dd-mmm-yyyy hh:nn:ss")
    Print #FileNum, "Time Taken     : " & Format(ProcessingEnd - ProcessingStart, "hh:mm:ss")
    Print #FileNum, ""
    Print #FileNum, "Total TC       : " & TotalTC
    Print #FileNum, "Passed         : " & PassedTC
    Print #FileNum, "Failed         : " & FailedTC
    Print #FileNum, ""

    If FailedTC > 0 Then
        Print #FileNum, "Failed TC(s)   :"
        Print #FileNum, FailedTCListText
    Else
        Print #FileNum, "Failed TC(s)   : None"
    End If

    Print #FileNum, "====================================================="

    Close #FileNum

End Sub
