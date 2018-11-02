Attribute VB_Name = "MXls_Z_Rg"
Option Explicit
Sub Z()
Z_RgNMoreBelow
End Sub

Function RgA1(A As Range) As Range
Set RgA1 = RgRC(A, 1, 1)
End Function

Function RgA1LasCell(A As Range) As Range
Dim L As Range, R, C
Set L = A.SpecialCells(xlCellTypeLastCell)
R = L.Row
C = L.Column
Set RgA1LasCell = WsRCRC(RgWs(A), A.Row, A.Column, R, C)
End Function

Function RgAdr$(A As Range)
RgAdr = "'" & RgWs(A).Name & "'!" & A.Address
End Function

Sub RgAsgRCRC(A As Range, OR1, OC1, OR2, OC2)
OR1 = A.Row
OR2 = OR1 + A.Rows.Count - 1
OC1 = A.Column
OC2 = OC1 + A.Columns.Count - 1
End Sub
Function CvRg(A) As Range
Set CvRg = A
End Function

Sub RgBdr(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub

Sub RgBdrAround(A As Range)
RgBdrL A
RgBdrR A
RgBdrTop A
RgBdrBottom A
End Sub

Sub RgBdrB(A As Range)
RgBdrL A
RgBdrR A
End Sub

Sub RgBdrBottom(A As Range)
RgBdr A, xlEdgeBottom
RgBdr A, xlEdgeTop
End Sub

Sub RgBdrInner(A As Range)
RgBdr A, xlInsideHorizontal
RgBdr A, xlInsideVertical
End Sub

Sub RgBdrInside(A As Range)
RgBdrInner A
End Sub

Sub RgBdrL(A As Range)
RgBdrLeft A
End Sub

Sub RgBdrLeft(A As Range)
RgBdr A, xlEdgeLeft
If A.Column > 1 Then
    RgBdr RgC(A, 0), xlEdgeRight
End If
End Sub

Sub RgBdrR(A As Range)
RgBdrRight A
End Sub

Sub RgBdrRight(A As Range)
RgBdr A, xlEdgeRight
If A.Column < MaxCol Then
    RgBdr RgC(A, A.Column + 1), xlEdgeLeft
End If
End Sub

Sub RgBdrTop(A As Range)
BdrSetLin A.Borders(xlEdgeTop)
If A.Row > 1 Then BdrSetLin A.Borders(xlEdgeBottom)
End Sub

Function RgC(A As Range, C) As Range
Set RgC = RgCC(A, C, C)
End Function

Function RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, RgNRow(A), C2)
End Function

Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function

Function RgDecBtmR(A As Range, Optional By% = 1) As Range
Set RgDecBtmR = RgRR(A, 1, A.Rows.Count + 1)
End Function

Sub RgeMgeV(A As Range)
Stop '?
End Sub

Function RgEntC(A As Range, C) As Range
Set RgEntC = RgC(A, C).EntireColumn
End Function

Sub RgFillCol(A As Range)
Dim Rg As Range
Dim Sq()
Sq = N_SqV(A.Rows.Count)
RgReSz(A, Sq).Value = Sq
End Sub

Sub RgFillRow(A As Range)
Dim Rg As Range
Dim Sq()
Sq = N_SqH(A.Rows.Count)
RgReSz(A, Sq).Value = Sq
End Sub

Function RgFstR(A As Range) As Range
Set RgFstR = RgR(A, 1)
End Function

Function RgFstC(A As Range) As Range
Set RgFstC = RgC(A, 1)
End Function

Function RgIncTopR(A As Range, Optional By% = 1) As Range
Set RgIncTopR = RgRR(A, 1 - By, A.Rows.Count)
End Function

Function RgIsHBar(A As Range) As Boolean
RgIsHBar = A.Rows.Count = 1
End Function

Function RgIsVBar(A As Range) As Boolean
RgIsVBar = A.Columns.Count = 1
End Function

Function RgLasCol%(A As Range)
RgLasCol = A.Column + A.Columns.Count - 1
End Function

Function RgLasHBar(A As Range) As Range
Set RgLasHBar = RgR(A, RgNRow(A))
End Function

Function RgLasRow&(A As Range)
RgLasRow = A.Row + A.Rows.Count - 1
End Function

Function RgLasVBar(A As Range) As Range
Set RgLasVBar = RgC(A, RgNCol(A))
End Function

Sub RgLnkWs(A As Range)
Dim R As Range
Dim WsNy$(): WsNy = WbWsNy(RgWb(A))
For Each R In A
    CellLnkWs R, WsNy
Next
End Sub

Function RgLo(A As Range) As ListObject
Dim Ws As Worksheet: Set Ws = RgWs(A)
Dim O As ListObject: Set O = Ws.ListObjects.Add(xlSrcRange, A, , xlYes)
RgBdrAround A
Set RgLo = O
End Function

Sub RgMge(A As Range)
A.MergeCells = True
A.HorizontalAlignment = XlHAlign.xlHAlignCenter
A.VerticalAlignment = XlVAlign.xlVAlignCenter
End Sub


Function RgNCol%(A As Range)
RgNCol = A.Columns.Count
End Function

Function RgNMoreBelow(A As Range, Optional N% = 1)
Set RgNMoreBelow = RgRR(A, 1, A.Rows.Count + N)
End Function

Function RgNMoreTop(A As Range, Optional N% = 1)
Dim O As Range
Set O = RgRR(A, 1 - N, A.Rows.Count)
Set RgNMoreTop = O
End Function

Function RgNRow&(A As Range)
RgNRow = A.Rows.Count
End Function

Function RgR(A As Range, R)
Set RgR = RgRR(A, R, R)
End Function

Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgReSz(A As Range, Sq) As Range
Set RgReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function

Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, RgNCol(A))
End Function

Function RgEntRR(A As Range, R1, R2) As Range
Set RgEntRR = RgRR(A, R1, R2).EntireRow
End Function

Function RgSq(A As Range)
If A.Columns.Count = 1 Then
    If A.Rows.Count = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = A.Value
        RgSq = O
        Exit Function
    End If
End If
RgSq = A.Value
End Function

Function RgVis(A As Range) As Range
XlsVis A.Application
Set RgVis = A
End Function

Function RgWb(A As Range) As Workbook
Set RgWb = WsWb(RgWs(A))
End Function

Function RgWs(A As Range) As Worksheet
Set RgWs = A.Parent
End Function


Private Sub Z_RgNMoreBelow()
Dim R As Range, Act As Range, Ws As Worksheet
Set Ws = NewWs
Set R = Ws.Range("A3:B5")
Set Act = RgNMoreTop(R, 1)
Debug.Print Act.Address
Stop
Debug.Print RgRR(R, 1, 2).Address
Stop
End Sub
