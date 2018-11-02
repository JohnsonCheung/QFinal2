Attribute VB_Name = "MXls_Z_Ws"
Option Explicit

Function WsC(A As Worksheet, C) As Range
Dim R As Range
Set R = A.Columns(C)
Set WsC = R.EntireColumn
End Function

Function WsCC(A As Worksheet, C1, C2) As Range
Set WsCC = WsRCC(A, 1, C1, C2).EntireColumn
End Function
Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Cells(1, 1)
End Function


Sub WsClrLo(A As Worksheet)
Dim Ay() As ListObject, J%
Ay = ItrInto(A.ListObjects, Ay)
For J = 0 To UB(Ay)
    Ay(J).Delete
Next
End Sub

Sub WsClsNoSav(A As Worksheet)
WbClsNoSav WsWb(A)
End Sub
Function CurWs() As Worksheet
Set CurWs = CurXls.ActiveSheet
End Function


Function WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(A, R1, C, R2, C)
End Function

Function WsDftLoNm$(A As Worksheet, Optional LoNm0$)
Dim LoNm$: LoNm = DftStr(LoNm0, "Table")
Dim J%
For J = 1 To 999
    If Not WsHasLoNm(A, LoNm) Then WsDftLoNm = LoNm: Exit Function
    LoNm = NmNxtSeqNm(LoNm)
Next
Stop
End Function

Function WsDlt(A As Workbook, WsIx) As Boolean
If WbHasWs(A, WsIx) Then WbWs(A, WsIx).Delete: Exit Function
WsDlt = True
End Function

Function WsDtaRg(A As Worksheet) As Range
Dim R, C
With WsLasCell(A)
   R = .Row
   C = .Column
End With
If R = 1 And C = 1 Then Exit Function
Set WsDtaRg = WsRCRC(A, 1, 1, R, C)
End Function

Function WsFstLo(A As Worksheet) As ListObject
Set WsFstLo = ItrFstItm(A.ListObjects)
End Function

Function WsHasLo(A As Worksheet, LoNm$) As Boolean
WsHasLo = ItrHasNm(A.ListObjects, LoNm)
End Function

Function WsHasLoNm(A As Worksheet, LoNm$) As Boolean
WsHasLoNm = WsHasLoNm(A, LoNm)
End Function

Function WsLasCell(A As Worksheet) As Range
Set WsLasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function WsLasCno%(A As Worksheet)
WsLasCno = WsLasCell(A).Column
End Function

Function WsLasRno%(A As Worksheet)
WsLasRno = WsLasCell(A).Row
End Function

Function WsLo(A As Worksheet, LoNm$) As ListObject
Set WsLo = A.ListObjects(LoNm)
End Function

Sub WsMinLo(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
ItrDo A.ListObjects, "LoMin"
End Sub

Function WsPtAy(A As Worksheet) As PivotTable()
Dim O() As PivotTable, Pt As PivotTable
For Each Pt In A.PivotTables
    PushObj O, Pt
Next
WsPtAy = O
End Function

Function WsPtNy(A As Worksheet) As String()
Dim Pt As PivotTable
For Each Pt In A.PivotTables
    PushI WsPtNy, Pt.Name
Next
End Function

Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function

Function WsRCC(A As Worksheet, R, C1, C2) As Range
Set WsRCC = WsRCRC(A, R, C1, R, C2)
End Function

Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function

Function WsRR(A As Worksheet, R1, R2) As Range
Set WsRR = A.Range(WsRC(A, R1, 1), WsRC(A, R2, 1)).EntireRow
End Function

Function WsSetLoNm(A As Worksheet, Nm$) As Worksheet
Dim Lo As ListObject
Set Lo = CvNothing(ItrFst(A.ListObjects))
If Not IsNothing(Lo) Then Lo.Name = "T_" & Nm
Set WsSetLoNm = A
End Function

Function WsSetNm(A As Worksheet, Nm$) As Worksheet
If Nm <> "" Then
    If Not WbHasWs(WsWb(A), Nm) Then A.Name = Nm
End If
Set WsSetNm = A
End Function

Function WsSq(A As Worksheet) As Variant()
WsSq = WsDtaRg(A).Value
End Function

Function WsVis(A As Worksheet) As Worksheet
XlsVis A.Application
Set WsVis = A
End Function

Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function
