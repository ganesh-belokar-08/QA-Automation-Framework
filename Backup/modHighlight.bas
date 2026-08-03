Attribute VB_Name = "modHighlight"
Option Explicit

Public Sub HighlightSheet(ws As Worksheet)

    Dim LastRowJ As Long
    Dim LastRowAA As Long

    Dim blueColor As Long
    Dim redColor As Long
    Dim yellowColor As Long
    Dim grayColor As Long

    On Error GoTo ErrHandler

    LastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    LastRowAA = ws.Cells(ws.Rows.Count, "AA").End(xlUp).Row

    If LastRowJ <= 1 Or LastRowAA <= 1 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ws.Cells.FormatConditions.Delete

    blueColor = RGB(97, 203, 243)
    redColor = RGB(255, 199, 206)
    yellowColor = RGB(255, 255, 153)
    grayColor = RGB(242, 242, 242)

    '==================================================
    ' BLUE
    '==================================================

    With ws.Range("J2:J" & LastRowJ)

        With .FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=OR(ISNUMBER(SEARCH(""."",J2)),ISNUMBER(SEARCH(""NONE"",J2)),J2=""20361"")")

            .Interior.Color = blueColor
            .StopIfTrue = True

        End With

    End With

    With ws.Range("AA2:AA" & LastRowAA)

        With .FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=OR(ISNUMBER(SEARCH(""."",AA2)),ISNUMBER(SEARCH(""NONE"",AA2)),AA2=""20361"")")

            .Interior.Color = blueColor
            .StopIfTrue = True

        End With

    End With

    '==================================================
    ' RED
    '==================================================

    With ws.Range("J2:J" & LastRowJ)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=AND(COUNTIF($AA:$AA,$J2)>0,SUMIF($J:$J,$J2,$M:$M)<>SUMIF($AA:$AA,$J2,$AC:$AC))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = redColor

    End With

    With ws.Range("K2:K" & LastRowJ)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=AND(COUNTIF($AA:$AA,$J2)>0,SUMIF($J:$J,$J2,$M:$M)<>SUMIF($AA:$AA,$J2,$AC:$AC))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = redColor

    End With

    With ws.Range("M2:M" & LastRowJ)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=AND(COUNTIF($AA:$AA,$J2)>0,SUMIF($J:$J,$J2,$M:$M)<>SUMIF($AA:$AA,$J2,$AC:$AC))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = redColor

    End With

    With ws.Range("AA2:AA" & LastRowAA)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=AND(COUNTIF($J:$J,$AA2)>0,SUMIF($AA:$AA,$AA2,$AC:$AC)<>SUMIF($J:$J,$AA2,$M:$M))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = redColor

    End With

    With ws.Range("AB2:AB" & LastRowAA)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=AND(COUNTIF($J:$J,$AA2)>0,SUMIF($AA:$AA,$AA2,$AC:$AC)<>SUMIF($J:$J,$AA2,$M:$M))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = redColor

    End With

    With ws.Range("AC2:AC" & LastRowAA)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=AND(COUNTIF($J:$J,$AA2)>0,SUMIF($AA:$AA,$AA2,$AC:$AC)<>SUMIF($J:$J,$AA2,$M:$M))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = redColor

    End With

    '==================================================
    ' YELLOW
    '==================================================

    With ws.Range("J2:J" & LastRowJ)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=ISERROR(MATCH(J2,$AA:$AA,0))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = yellowColor

    End With

    With ws.Range("AA2:AA" & LastRowAA)

        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:="=ISERROR(MATCH(AA2,$J:$J,0))"

        .FormatConditions(.FormatConditions.Count).Interior.Color = yellowColor

    End With

    '==================================================
    ' GRAY
    '==================================================

    With ws.Range("J2:J" & LastRowJ)

        .FormatConditions.Add _
            Type:=xlCellValue, _
            Operator:=xlNotEqual, _
            Formula1:="="""""

        .FormatConditions(.FormatConditions.Count).Interior.Color = grayColor

    End With

    With ws.Range("AA2:AA" & LastRowAA)

        .FormatConditions.Add _
            Type:=xlCellValue, _
            Operator:=xlNotEqual, _
            Formula1:="="""""

        .FormatConditions(.FormatConditions.Count).Interior.Color = grayColor

    End With

ExitHandler:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:

    Resume ExitHandler

End Sub

