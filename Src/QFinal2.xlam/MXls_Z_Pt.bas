Attribute VB_Name = "MXls_Z_Pt"
Option Explicit
Function PtCpyToLo(A As PivotTable, At As Range) As ListObject
Dim R1, R2, C1, C2, NC, NR
    R1 = A.RowRange.Row
    C1 = A.RowRange.Column
    R2 = RgLasRow(A.DataBodyRange)
    C2 = RgLasCol(A.DataBodyRange)
    NC = C2 - C1 + 1
    NR = R2 - C1 + 1
WsRCRC(PtWs(A), R1, C1, R2, C2).Copy
At.PasteSpecial xlPasteValues

Set PtCpyToLo = RgLo(RgRCRC(At, 1, 1, NR, NC))
End Function

Sub PtFFSetOri(A As PivotTable, FF, Ori As XlPivotFieldOrientation)
Dim F, J%, T
T = Array(False, False, False, False, False, False, False, False, False, False, False, False)
J = 1
For Each F In AyNz(SslSy(FF))
    With PtPf(A, F)
        .Orientation = Ori
        .Position = J
        If Ori = xlColumnField Or Ori = xlRowField Then
            .Subtotals = T
        End If
    End With
    J = J + 1
Next
End Sub

Private Sub PtFmt()

End Sub

Function PtPf(A As PivotTable, F) As PivotField
Set PtPf = A.PivotFields(F)
End Function

Function PtRowFldEntCol(A As PivotTable, F) As Range
Set PtRowFldEntCol = RgR(PtPf(A, F).DataRange, 1).EntireColumn
End Function

Sub PtSetRowssColWdt(A As PivotTable, Rowss$, ColWdt As Byte)
If ColWdt <= 1 Then Stop
Dim F
For Each F In AyNz(SslSy(Rowss))
    PtRowFldEntCol(A, F).ColumnWidth = ColWdt
Next
End Sub

Sub PtSetRowssOutLin(A As PivotTable, Rowss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim F
For Each F In AyNz(SslSy(Rowss))
    PtRowFldEntCol(A, F).OutlineLevel = Lvl
Next
End Sub

Sub PtSetRowssRepeatLbl(A As PivotTable, Rowss$)
Dim F
For Each F In AyNz(SslSy(Rowss))
    PtPf(A, F).RepeatLabels = True
Next
End Sub

Function PtVis(A As PivotTable) As PivotTable
A.Application.Visible = True
Set PtVis = A
End Function

Function PtWb(A As PivotTable) As Workbook
Set PtWb = WsWb(PtWs(A))
End Function

Function PtWs(A As PivotTable) As Worksheet
Set PtWs = A.Parent
End Function
