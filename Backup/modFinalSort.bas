Attribute VB_Name = "modFinalSort"
Option Explicit

Public Sub FinalSort(ws As Worksheet)

    Dim LastRowJ As Long
    Dim LastRowAA As Long

    On Error GoTo ErrHandler

    '----------------------------------------
    ' Find Last Rows
    '----------------------------------------

    LastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    LastRowAA = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    If LastRowJ <= 1 Or LastRowAA <= 1 Then Exit Sub

    '========================================
    ' FINAL SORT LEFT SIDE (J:M)
    '========================================

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
            Key:=ws.Range("J2:J" & LastRowJ), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange ws.Range("J2:M" & LastRowJ)

        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

    '========================================
    ' FINAL SORT RIGHT SIDE (AA:AC)
    '========================================

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
            Key:=ws.Range("AA2:AA" & LastRowAA), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange ws.Range("AA2:AC" & LastRowAA)

        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

ExitHandler:
    Exit Sub

ErrHandler:
    Resume ExitHandler

End Sub

