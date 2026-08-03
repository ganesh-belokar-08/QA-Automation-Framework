Attribute VB_Name = "modPrioritySort"
Option Explicit

'=========================================================
' Assign Priority Based on Cell Color
'=========================================================

Public Sub AssignPriority(ws As Worksheet)

    Dim LastRowJ As Long
    Dim LastRowAA As Long
    Dim r As Long

    LastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    LastRowAA = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    'Hide helper columns
    ws.Columns("N").Hidden = True
    ws.Columns("AD").Hidden = True

    '-----------------------------
    'Left Block (J:M)
    '-----------------------------
    For r = 2 To LastRowJ

        Select Case ws.Cells(r, "J").DisplayFormat.Interior.Color

            Case RGB(242, 242, 242)      'Gray
                ws.Cells(r, "N") = 1

            Case RGB(255, 199, 206)      'Red
                ws.Cells(r, "N") = 2

            Case RGB(255, 255, 153)      'Yellow
                ws.Cells(r, "N") = 3

            Case RGB(97, 203, 243)       'Blue
                ws.Cells(r, "N") = 4

            Case Else
                ws.Cells(r, "N") = 5

        End Select

    Next r

    '-----------------------------
    'Right Block (AA:AC)
    '-----------------------------
    For r = 2 To LastRowAA

        Select Case ws.Cells(r, "AA").DisplayFormat.Interior.Color

            Case RGB(242, 242, 242)      'Gray
                ws.Cells(r, "AD") = 1

            Case RGB(255, 199, 206)      'Red
                ws.Cells(r, "AD") = 2

            Case RGB(255, 255, 153)      'Yellow
                ws.Cells(r, "AD") = 3

            Case RGB(97, 203, 243)       'Blue
                ws.Cells(r, "AD") = 4

            Case Else
                ws.Cells(r, "AD") = 5

        End Select

    Next r

End Sub

'=========================================================
' Sort By Priority then Product
'=========================================================

Public Sub SortByPriority(ws As Worksheet)

    Dim LastRowJ As Long
    Dim LastRowAA As Long

    LastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    LastRowAA = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    '----------------------------------------
    'Left Block
    '----------------------------------------
    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
        Key:=ws.Range("N2:N" & LastRowJ), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending

        .SortFields.Add _
        Key:=ws.Range("J2:J" & LastRowJ), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending

        .SetRange ws.Range("J1:N" & LastRowJ)

        .Header = xlYes
        .Apply

    End With

    '----------------------------------------
    'Right Block
    '----------------------------------------
    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
        Key:=ws.Range("AD2:AD" & LastRowAA), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending

        .SortFields.Add _
        Key:=ws.Range("AA2:AA" & LastRowAA), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending

        .SetRange ws.Range("AA1:AD" & LastRowAA)

        .Header = xlYes
        .Apply

    End With

End Sub

'=========================================================
' Delete Helper Values
'=========================================================

Public Sub DeletePriority(ws As Worksheet)

    ws.Columns("N").ClearContents
    ws.Columns("AD").ClearContents

    ws.Columns("N").Hidden = False
    ws.Columns("AD").Hidden = False

End Sub

