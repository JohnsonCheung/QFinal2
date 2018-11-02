Attribute VB_Name = "MXls_Z_Wb"
Option Explicit
Sub Z()
Z_WbSetFcsv
Z_WbTxtCnCnt
End Sub
Function WbOleWcAy(A As Workbook) As OLEDBConnection()
'Dim O() As OLEDBConnection, Wc As WorkbookConnection
'For Each Wc In A.Connections
'    PushObjNonNothingObj O, Wc.OLEDBConnection
'Next
'WbOleWcAy = O
WbOleWcAy = OyRmvNothing(ItrpAyInto(A.Connections, "OLEDBConnection", WbOleWcAy))
End Function

Function WbTxtCnStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
If IsNothing(T) Then Exit Function
WbTxtCnStr = T.Connection
End Function
Function CurWb() As Workbook
Set CurWb = CurXls.ActiveWorkbook
End Function


Sub WbVdtOupNy(A As Workbook, OupNy$())
Dim O$(), N$, B$(), WsCdNy$()
WsCdNy = WbWsCdNy(A)
O = AyMinus(AyAddPfx(OupNy, "WsO"), WsCdNy)
If Sz(O) > 0 Then
    N = "OupNy":  B = OupNy:  GoSub Dmp
    N = "WbCdNy": B = WsCdNy: GoSub Dmp
    N = "Mssing": B = O:      GoSub Dmp
    Stop
    Exit Sub
End If
Exit Sub
Dmp:
Debug.Print UnderLin(N)
Debug.Print N
Debug.Print UnderLin(N)
AyDmp B
Return
End Sub

Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function

Function WbWcNy(A As Workbook) As String()
WbWcNy = ItrNy(A.Connections)
End Function

Function WbWcSy_zOle(A As Workbook) As String()
WbWcSy_zOle = OyPrpSy(WbOleWcAy(A), "Connection")
End Function

Function WbWs(A As Workbook, WsNm) As Worksheet
Set WbWs = A.Sheets(WsNm)
End Function

Function WbWsCd(A As Workbook, WsCdNm$) As Worksheet
Set WbWsCd = ItrFstPrpEqV(A.Sheets, "CodeName", WsCdNm)
End Function

Function WbWsCdNy(A As Workbook) As String()
WbWsCdNy = ItrPrpSy(A.Sheets, "CodeName")
End Function

Function WbWsNy(A As Workbook) As String()
WbWsNy = ItrNy(A.Sheets)
End Function

Private Sub Z_WbSetFcsv()
Dim Wb As Workbook
'Set Wb = FxWb(VbeMthFx)
Debug.Print WbTxtCnStr(Wb)
WbSetFcsv Wb, "C:\ABC.CSV"
Ass WbTxtCnStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub

Private Sub Z_WbTxtCnCnt()
Dim O As Workbook: 'Set O = FxWb(VbeMthFx)
Ass WbTxtCnCnt(O) = 1
O.Close
End Sub

Private Sub ZZ_WbLoAy()
'D OyNy(WbLoAy(TpWb))
End Sub

Private Sub ZZ_WbTLoAy()
'D OyNy(WbTLoAy(TpWb))
End Sub

Sub ZZ_WbWcSy()
'D WbWcSy_zOle(FxWb(TpFx))
End Sub

Function WbNewA1(A As Workbook, Optional WsNm$) As Range
Set WbNewA1 = WsA1(WbAddWs(A, WsNm))
End Function

Function WbAddDbt(A As Workbook, Db As Database, T$, Optional UseWc As Boolean) As Workbook
'Set WbAddDbt = LoWb(DbtAtLo(Db, T, WsA1(A, T), UseWc))
End Function

Function WbAddDbtt(A As Workbook, Db As Database, TT, Optional UseWc As Boolean) As Workbook
AyDoPPXP CvNy(TT), "WbAddDbt", A, Db, UseWc
Set WbAddDbtt = A
End Function

Function WbAddDt(A As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(A, Dt.DtNm)
DrsLo DtDrs(Dt), WsA1(O)
Set WbAddDt = O
End Function

Function WbAddWc(A As Workbook, Fb$, Nm$) As WorkbookConnection
Set WbAddWc = A.Connections.Add2(Nm, Nm, FbWcStr(Fb), Nm, XlCmdType.xlCmdTable)
End Function

Function WbAddWs(A As Workbook, Optional WsNm$, Optional AtBeg As Boolean, Optional AtEnd As Boolean, Optional BefWsNm$, Optional AftWsNm$) As Worksheet
Dim O As Worksheet
WbDltWs A, WsNm
Select Case True
Case AtBeg:         Set O = A.Sheets.Add(WbFstWs(A))
Case AtEnd:         Set O = A.Sheets.Add(WbLasWs(A))
Case BefWsNm <> "": Set O = A.Sheets.Add(A.Sheets(BefWsNm))
Case AftWsNm <> "": Set O = A.Sheets.Add(, A.Sheets(AftWsNm))
Case Else:          Set O = A.Sheets.Add
End Select
Set WbAddWs = WsSetNm(O, WsNm)
End Function

Function WbCdNmWs(A As Workbook, CdNm$) As Worksheet
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.CodeName = CdNm Then Set WbCdNmWs = Ws: Exit Function
Next
End Function

Sub WbClsNoSav(A As Workbook)
A.Close False
End Sub

Function WbCn_TxtCn(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set WbCn_TxtCn = A.TextConnection
End Function

Sub WbDltWc(A As Workbook)
ItrDo A.Connections, "WcDlt"
End Sub

Sub WbDltWs(A As Workbook, WsNm)
If WbHasWs(A, WsNm) Then
    A.Application.DisplayAlerts = False
    WbWs(A, WsNm).Delete
    A.Application.DisplayAlerts = True
End If
End Sub

Sub WbFmt(A As Workbook, WbFmtrAv())
Dim J%
For J = 0 To UB(WbFmtrAv)
    Run WbFmtrAv(J), A
Next
WbMax(WbVis(A)).Save
End Sub

Sub WbFmtAllLo(A As Workbook)
FmtSpec_Imp
'AyBrwThw FmtSpec_ErLy
AyDoXP WbLoAy(A), "LoFmt", FmtSpec_Ly
End Sub

Function WbFstWs(A As Workbook) As Worksheet
Set WbFstWs = A.Sheets(1)
End Function

Function WbFx$(A As Workbook)
Dim F$
F = A.FullName
If F = A.Name Then Exit Function
WbFx = F
End Function

Function WbHasWs(A As Workbook, WsNm) As Boolean
WbHasWs = ItrHasNm(A.Sheets, WsNm)
End Function

Function WbLasWs(A As Workbook) As Worksheet
Set WbLasWs = A.Sheets(A.Sheets.Count)
End Function

Function WbLo(A As Workbook, LoNm$) As ListObject
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If WsHasLo(Ws, LoNm) Then Set WbLo = Ws.ListObjects(LoNm): Exit Function
Next
End Function

Function WbLoAy(Tp As Workbook) As ListObject()
Dim Ws As Worksheet
For Each Ws In Tp.Sheets
    'PushItr WbLoAy, Ws.ListObjects
Next
End Function

Function WbMainLo(A As Workbook) As ListObject
Dim O As Worksheet, Lo As ListObject
Set O = WbMainWs(A):              If IsNothing(O) Then Exit Function
Set WbMainLo = WsLo(O, "T_Main")
End Function

Function WbMainQt(A As Workbook) As QueryTable
Dim Lo As ListObject
Set Lo = WbMainLo(A): If IsNothing(A) Then Exit Function
Set WbMainQt = Lo.QueryTable
End Function

Function WbMainWs(A As Workbook) As Worksheet
Set WbMainWs = WbWsCd(A, "WsOMain")
End Function

Function WbMax(A As Workbook) As Workbook
A.Application.WindowState = xlMaximized
Set WbMax = A
End Function

Function WbMinLo(A As Workbook) As Workbook
ItrDo A.Sheets, "WsMinLo"
Set WbMinLo = A
End Function

Function WbMthMdDrs(A As Workbook) As Drs
Dim Lo As ListObject
Set Lo = WbLo(A, "T_MthMd")
If IsNothing(Lo) Then Exit Function
If Not IsEqAy(LoFny(Lo), SslSy("Mth Md")) Then Stop
Set WbMthMdDrs = LoDrs(Lo)
End Function

Function WbOupLoAy(A As Workbook) As ListObject()
WbOupLoAy = OyWhNmHasPfx(WbLoAy(A), "T_")
End Function

Function WbPtAy(A As Workbook) As PivotTable()
Dim O() As PivotTable, Ws As Worksheet
For Each Ws In A.Sheets
    PushObjAy O, WsPtAy(Ws)
Next
WbPtAy = O
End Function

Function WbPtNy(A As Workbook) As String()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushIAy WbPtNy, WsPtNy(Ws)
Next
End Function

Sub WbQuit(A As Workbook)
XlsQuit A.Application
End Sub

Function WbSav(A As Workbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.Save
A.Application.DisplayAlerts = Y
Set WbSav = A
End Function

Function WbSavAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.SaveAs Fx, Fmt
A.Application.DisplayAlerts = Y
Set WbSavAs = A
End Function

Sub WbSetFcsv(A As Workbook, Fcsv$)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub

Function WbTLoAy(A As Workbook) As ListObject()
WbTLoAy = OyWhNmHasPfx(WbLoAy(A), "T_")
End Function

Function WbTxtCn(A As Workbook) As TextConnection
Dim N%: N = WbTxtCnCnt(A)
If N <> 1 Then
    Stop
    Exit Function
End If
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then
        Set WbTxtCn = C.TextConnection
        Exit Function
    End If
Next
ErImposs
End Function

Function WbTxtCnCnt%(A As Workbook)
Dim C As WorkbookConnection, Cnt%
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then Cnt = Cnt + 1
Next
WbTxtCnCnt = Cnt
End Function
