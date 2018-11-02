Attribute VB_Name = "MXls_Z_Lo"
Option Explicit
Sub Z()
Z_LoPt
Z_LoReset
End Sub
Function TnLoNm$(TblNm)
TnLoNm = "T_" & RmvFstNonLetter(TblNm)
End Function

Private Sub Z_LoPt()
Dim At As Range, Lo As ListObject
Set Lo = SampleLo
Set At = RgVis(WsA1(WbAddWs(LoWb(Lo))))
PtVis LoPt(Lo, At, "A B", "C D", "F", "E")
Stop
End Sub

Function SqLo(A, At As Range) As ListObject
Set SqLo = RgLo(SqRg(A, At))
End Function



Function CvLo(A) As ListObject
Set CvLo = A
End Function

Private Sub Z_LoReset()
Dim Wb As Workbook, LoAy() As ListObject
Set Wb = FxWb("C:\users\user\desktop\a.xlsx")
WbVis Wb
LoAy = WbLoAy(Wb)
'LoReset LoAy(0)
End Sub

Sub ZZ_LoKeepFstCol()
LoKeepFstCol LoVis(SampleLo)
End Sub

Sub LoAutoFit(A As ListObject)
Dim C As Range: Set C = LoAllEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = RgEntC(C, J)
   If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

Sub Z_LoAutoFit()
Dim Ws As Worksheet: Set Ws = NewWs
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
LoAutoFit LoCrt(Ws)
WsClsNoSav Ws
End Sub

Function LoBdrAround(A As ListObject)
Dim R As Range
Set R = RgNMoreTop(A.DataBodyRange)
If A.ShowTotals Then Set R = RgNMoreBelow(R)
RgBdrAround R
End Function

Sub LoBrw(A As ListObject)
DrsBrw LoDrs(A)
End Sub

Sub Z_LoBrw()
Dim O As ListObject: Set O = SampleLo
LoBrw O
Stop
End Sub

Function LoCol(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R As Range
Set R = A.ListColumns(C).DataBodyRange
If Not InclTot And Not InclHdr Then
    Set LoCol = R
    Exit Function
End If

If InclTot Then Set LoCol = RgDecBtmR(R, 1)
If InclHdr Then Set LoCol = RgIncTopR(R, 1)
End Function

Function LoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = LoR1(A, InclHdr)
R2 = LoR2(A, InclTot)
mC1 = LoWsCno(A, C1)
mC2 = LoWsCno(A, C2)
Set LoCC = WsRCRC(LoWs(A), R1, mC1, R2, mC2)
End Function

Sub LoCol_LnkWs(A As ListObject, C)
RgLnkWs LoCol_Rg(A, C)
End Sub

Function LoCol_Rg(A As ListObject, C) As Range
Set LoCol_Rg = A.ListColumns(C).Range
End Function


Function LoCrt(A As Worksheet, Optional LoNm$) As ListObject
Dim R As Range: Set R = WsDtaRg(A)
If IsNothing(R) Then Exit Function
Dim O As ListObject: Set O = A.ListObjects.Add(xlSrcRange, WsDtaRg(A), , xlYes)
If LoNm <> "" Then O.Name = LoNm
LoAutoFit O
Set LoCrt = O
End Function


Sub LoDlt(A As ListObject)
Dim R As Range, R1, C1, R2, C2, Ws As Worksheet
Set Ws = LoWs(A)
Set R = RgNMoreBelow(RgNMoreTop(A.DataBodyRange))
RgAsgRCRC R, R1, C1, R2, C2
A.QueryTable.Delete
WsRCRC(Ws, R1, C1, R2, C2).ClearContents
End Sub

Function LoDrs(A As ListObject) As Drs
Set LoDrs = Drs(LoFny(A), LoDry(A))
End Function

Function LoDry(A As ListObject) As Variant()
LoDry = SqDry(LoSq(A))
End Function

Function LoDrySel(A As ListObject, FF) As Variant() _
' Return as many column as fields in [FF] from Lo[A]
Dim IxAy&(), Dry(): GoSub X_IxAy_Dry
Dim Dr
For Each Dr In AyNz(Dry)
    PushI LoDrySel, DrSel(Dr, IxAy)
Next
Exit Function
X_IxAy_Dry:
    Dim Fny$()
    Fny = LoFny(A)
    Dry = LoDry(A)
    IxAy = AyIxAy(Fny, SslSy(FF))
    Return
End Function

Function LoDtaAdr$(A As ListObject)
LoDtaAdr = RgAdr(A.DataBodyRange)
End Function

Sub LoEnsNRow(A As ListObject, NRow&)
LoMin A
Exit Sub
If NRow > 1 Then
    Debug.Print A.InsertRowRange.Address
    Stop
End If
End Sub

Function LoAllCol(A As ListObject) As Range
Set LoAllCol = LoCC(A, 1, LoNCol(A))
End Function

Function LoAllEntCol(A As ListObject) As Range
Set LoAllEntCol = LoAllCol(A).EntireColumn
End Function

Function LoFbtStr$(A As ListObject)
LoFbtStr = QtFbtStr(A.QueryTable)
End Function

Function LoFny(A As ListObject) As String()
LoFny = ItrNy(A.ListColumns)
End Function

Function LoHasFny(A As ListObject, Fny$()) As Boolean
LoHasFny = AyHasAy(LoFny(A), Fny)
End Function

Function LoHasNoDta(A As ListObject) As Boolean
LoHasNoDta = IsNothing(A.DataBodyRange)
End Function

Function LoHdrCell(A As ListObject, FldNm) As Range
Dim Rg As Range: Set Rg = A.ListColumns(FldNm).Range
Set LoHdrCell = RgRC(Rg, 1, 1)
End Function

Sub LoKeepFstCol(A As ListObject)
Dim J%
For J = A.ListColumns.Count To 2 Step -1
    A.ListColumns(J).Delete
Next
End Sub

Sub LoKeepFstRow(A As ListObject)
Dim J%
For J = A.ListRows.Count To 2 Step -1
    A.ListRows(J).Delete
Next
End Sub

Sub LoMin(A As ListObject)
Dim R1 As Range, R2 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    Set R2 = RgRR(R1, 2, R1.Rows.Count)
    R2.Delete
End If
End Sub

Function LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Function

Function LoNmTblNm$(A)
If Not HasPfx(A, "T_") Then Stop
LoNmTblNm = "@" & Mid(A, 3)
End Function

Function LoPc(A As ListObject) As PivotCache
Dim O As PivotCache
Set O = LoWb(A).PivotCaches.Create(xlDatabase, A.Name, 6)
O.MissingItemsLimit = xlMissingItemsNone
Set LoPc = O
End Function

Function LoPt(A As ListObject, At As Range, Rowss$, Dtass$, Optional Colss$, Optional Pagss$) As PivotTable
If LoWb(A).FullName <> RgWb(At).FullName Then Stop: Exit Function
Dim O As PivotTable
Set O = LoPc(A).CreatePivotTable(TableDestination:=At, TableName:=LoPtNm(A))
With O
    .ShowDrillIndicators = False
    .InGridDropZones = False
    .RowAxisLayout xlTabularRow
End With
O.NullString = ""
PtFFSetOri O, Rowss, xlRowField
PtFFSetOri O, Colss, xlColumnField
PtFFSetOri O, Pagss, xlPageField
PtFFSetOri O, Dtass, xlDataField
Set LoPt = O
End Function

Function LoPtNm$(A As ListObject)
If Left(A.Name, 2) <> "T_" Then Stop
Dim O$: O = "P_" & Mid(A.Name, 3)
LoPtNm = AyNxtNm(WbPtNy(LoWb(A)), O)
End Function

Function LoQt(A As ListObject) As QueryTable
On Error Resume Next
Set LoQt = A.QueryTable
End Function

Function LoR1&(A As ListObject, Optional InclHdr As Boolean)
If LoHasNoDta(A) Then
   LoR1 = A.ListColumns(1).Range.Row + 1
   Exit Function
End If
LoR1 = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function LoR2&(A As ListObject, Optional InclTot As Boolean)
If LoHasNoDta(A) Then
   LoR2 = LoR1(A)
   Exit Function
End If
LoR2 = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Sub LoReset(A As ListObject, WFb$)
'When LoRfh, if the fields of Db's table has been reorder, the Lo will not follow the order
'Delete the Lo, add back the Wc then WcAt to reset the Lo
'TblNm : from Lo.Name = T_XXX is the key to get the table name.
'Fb    : use WFb
Dim LoNm$, T$, At As Range, Wb As Workbook
Set Wb = LoWb(A)
Set At = RgRC(A.DataBodyRange, 0, 1)
LoNm = A.Name
T = LoNmTblNm(LoNm)
LoDlt A
WcAt WbAddWc(Wb, WFb, T), At
End Sub


Function LoSel(A As ListObject, FF) As Drs
Dim Fny$(): Fny = LinTermAy(FF)
Set LoSel = Drs(Fny, SqSel(LoSq(A), AyIxAy(LoFny(A), Fny)))
End Function

Function LoEntCol(A As ListObject, C) As Range
Set LoEntCol = LoCol(A, C).EntireColumn
End Function

Function LoSq(A As ListObject)
LoSq = A.DataBodyRange.Value
End Function

Function LoColSy(A As ListObject, C) As String()
LoColSy = SqColSy(A.ListColumns(C).DataBodyRange.Value, 1)
End Function

Function LoVis(A As ListObject) As ListObject
XlsVis A.Application
Set LoVis = A
End Function

Function LoWb(A As ListObject) As Workbook
Set LoWb = WsWb(LoWs(A))
End Function

Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function

Function LoWsCno%(A As ListObject, Col)
LoWsCno = A.ListColumns(Col).Range.Column
End Function

Function LoEnsCol(A As ListObject, C) As Boolean
Const CSub$ = "LoEnsCol"
If LoHasCol(A, C) Then
    Warn CSub, "[Lo] does not have [Col]", A.Name, C
    Exit Function
End If
LoEnsCol = True
End Function
Function LoHasCol(A As ListObject, C) As Boolean
LoHasCol = ItrHasNm(A.ListColumns, C)
End Function

Function LoSetNm(A As ListObject, LoNm$) As ListObject
If LoNm <> "" Then
    If Not WsHasLo(A, LoNm) Then
        A.Name = LoNm
    End If
End If
Set LoSetNm = A
End Function

