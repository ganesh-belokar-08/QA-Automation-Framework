Attribute VB_Name = "modController"
Option Explicit

'=========================================================
' Run Complete QA Automation Framework
'=========================================================

Public Sub RunFramework(TargetWB As Workbook)

    Dim ws As Worksheet
    Dim IsPass As Boolean
    Dim FailReason As String

    ProcessingStart = Now

    TotalTC = 0
    PassedTC = 0
    FailedTC = 0
    FailedTCList = ""
    FailedTCListText = ""

    For Each ws In TargetWB.Worksheets

        If IsTCSheet(ws) Then

            TotalTC = TotalTC + 1

            Application.StatusBar = _
                "Processing " & TotalTC & " : " & ws.Name

            '--------------------------------
            ' Step 1 - Highlight Differences
            '--------------------------------
            HighlightSheet ws

            '--------------------------------
            ' Step 2 - Assign Priority
            '--------------------------------
            AssignPriority ws

            '--------------------------------
            ' Step 3 - Check Pass / Fail
            '--------------------------------
            IsPass = CheckSheetResult(ws, FailReason)

            If IsPass Then
                PassedTC = PassedTC + 1
            Else
                FailedTC = FailedTC + 1

                If FailedTCList = "" Then
                    FailedTCList = ws.Name & " (" & FailReason & ")"
                Else
                    FailedTCList = FailedTCList & ", " & ws.Name & " (" & FailReason & ")"
                End If

                FailedTCListText = FailedTCListText & ws.Name & " (" & FailReason & ")" & vbCrLf
            End If

            '--------------------------------
            ' Step 4 - Sort by Priority
            '--------------------------------
            SortByPriority ws

            '--------------------------------
            ' Step 5 - Remove Helper Columns
            '--------------------------------
            DeletePriority ws

        End If

    Next ws

    ProcessingEnd = Now

    Application.StatusBar = False

End Sub

'=========================================================
' Check Pass / Fail for a Sheet
' FAIL = data missing in J or M (left block)
'        OR data missing in AA or AC (right block)
' FailReason returns which column(s) are empty
'=========================================================

Private Function CheckSheetResult(ws As Worksheet, ByRef FailReason As String) As Boolean

    Dim LastRowJ As Long
    Dim LastRowM As Long
    Dim LastRowAA As Long
    Dim LastRowAC As Long
    Dim Reasons As String

    LastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    LastRowM = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    LastRowAA = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row
    LastRowAC = ws.Cells(ws.Rows.Count, "AC").End(xlUp).Row

    Reasons = ""

    If LastRowJ < 2 Then Reasons = Reasons & "J empty, "
    If LastRowM < 2 Then Reasons = Reasons & "M empty, "
    If LastRowAA < 2 Then Reasons = Reasons & "AA empty, "
    If LastRowAC < 2 Then Reasons = Reasons & "AC empty, "

    If Reasons = "" Then
        CheckSheetResult = True    'PASS
        FailReason = ""
    Else
        CheckSheetResult = False   'FAIL
        FailReason = Left(Reasons, Len(Reasons) - 2)   'trim trailing ", "
    End If

End Function

