Attribute VB_Name = "modSortColor"
Option Explicit

Public Sub SortByColor(ws As Worksheet)

    Dim rng1 As Range
    Dim rng2 As Range
    Dim productCol1 As Long
    Dim productCol2 As Long
    Dim LastRowJ As Long
    Dim LastRowAA As Long

    On Error GoTo ErrHandler

    LastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    LastRowAA = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    If LastRowJ <= 1 Or LastRowAA <= 1 Then Exit Sub

    Set rng1 = ws.Range("J1:M" & LastRowJ)
    Set rng2 = ws.Range("AA1:AC" & LastRowAA)

    productCol1 = rng1.Columns(1).Column
    productCol2 = rng2.Columns(1).Column

    '==================================================
    ' SORT LEFT SIDE (J:M)
    '==================================================

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol1), ws.Cells(LastRowJ, productCol1)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(242, 242, 242)

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol1), ws.Cells(LastRowJ, productCol1)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(255, 199, 206)

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol1), ws.Cells(LastRowJ, productCol1)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(255, 255, 153)

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol1), ws.Cells(LastRowJ, productCol1)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(97, 203, 243)

        .SetRange rng1
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

    '==================================================
    ' SORT RIGHT SIDE (AA:AC)
    '==================================================

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol2), ws.Cells(LastRowAA, productCol2)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(242, 242, 242)

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol2), ws.Cells(LastRowAA, productCol2)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(255, 199, 206)

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol2), ws.Cells(LastRowAA, productCol2)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(255, 255, 153)

        .SortFields.Add( _
            ws.Range(ws.Cells(1, productCol2), ws.Cells(LastRowAA, productCol2)), _
            xlSortOnCellColor, xlAscending).SortOnValue.Color = RGB(97, 203, 243)

        .SetRange rng2
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

ExitHandler:
    Exit Sub

ErrHandler:
    Resume ExitHandler

End Sub

