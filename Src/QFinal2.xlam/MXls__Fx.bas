Attribute VB_Name = "MXls__Fx"
Option Explicit
Function FxBrw(A$)
WbVis FxWb(A)
End Function


Sub FxCrt(A)
WbSavAs(NewWb, A).Close
End Sub

Function FxDftWsNm$(A, WsNm0$)
If WsNm0 = "" Then
    FxDftWsNm = FxFstWsNm(A)
    Exit Function
End If
FxDftWsNm = WsNm0
End Function

Function FxDftWsNy(A, WsNy0) As String()
Dim WsNy$(): WsNy = CvNy(WsNy0)
If Sz(WsNy) = 0 Then
   FxDftWsNy = FxWsNy(A)
   Exit Function
End If
FxDftWsNy = WsNy
End Function

Function FxEns$(A$)
If Not IsFfnExist(A) Then FxCrt A
FxEns = A
End Function

Function FxFny(A$, Optional WsNm$ = "Sheet1") As String()
FxFny = ItrNy(FxCat(A).Tables(WsNm & "$").Columns)
End Function

Function FxFstWsNm$(A)
FxFstWsNm = RmvLasChr(CvCatTbl(ItrFst(FxCat(A).Tables)).Name)
End Function

Sub FxMinLo(A)
WbClsNoSav WbSav(WbMinLo(FxWb(A)))
End Sub

Function FxMthMdDrs(A$) As Drs
Dim Wb As Workbook
Set Wb = FxWb(A)
Set FxMthMdDrs = WbMthMdDrs(Wb)
Wb.Close False
End Function

Function FxOleCnStr$(A)
FxOleCnStr = "OLEDb;" & FxAdoCnStr(A)
End Function

Sub FxOpn(A)
If Not FfnIsExist(A) Then
    MsgBox "File not found: " & vbCrLf & vbCrLf & A
    Exit Sub
End If
Dim C$
C = FmtQQ("Excel ""?""", A)
Shell C, vbMaximizedFocus
End Sub

Sub FxRmvWsIfExist(A, WsNm)
If FxHasWs(A, WsNm) Then
   Dim B As Workbook: Set B = FxWb(A)
   WbWs(B, WsNm).Delete
   WbSav B
   WbClsNoSav B
End If
End Sub

Function FxSqlDrs(A, Sql) As Drs
Set FxSqlDrs = RsDrs(FxCn(A).Execute(Sql))
End Function

Sub FxSqlRun(A, Sql)
FxCn(A).Execute Sql
End Sub

Function FxTmpDb(Fx$, Optional WsNy0) As Database
Dim O As Database
   Set O = TmpDb
   DbLnkFx O, Fx, WsNy0
Set FxTmpDb = O
End Function

Private Sub Z_FxTmpDb()
Dim Db As Database: Set Db = FxTmpDb("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
AyDmp DbTny(Db)
Db.Close
End Sub

Function FxWb(A) As Workbook
Set FxWb = Xls.Workbooks.Open(A)
End Function

Function FxWs(A, Optional WsNm$ = "Data") As Worksheet
Set FxWs = WbWs(FxWb(A), WsNm)
End Function

Function FxWsARs(A, W, Optional F = 0) As ADODB.Recordset
Dim T$
T = W & "$"
If F = 0 Then
    Q = QSel_Fm(T)
Else
    Q = QSel_FF_Fm(F, T)
End If
Set FxWsARs = CnqRs(FxCn(A), Q)
End Function

Function FxWsCdNy(A) As String()
Dim Wb As Workbook
Set Wb = FxWb(A)
FxWsCdNy = WbWsCdNy(Wb)
Wb.Close False
End Function
Function FxWsChk(A$, Optional WsNy0 = "Sheet1", Optional FxKind$ = "Excel file") As String()
If FfnNotExist(A) Then FxWsChk = FfnNotExistMsg(A, FxKind): Exit Function
If FxHasWs(A, WsNy0) Then Exit Function
Dim M$
M = FmtQQ("[?] in [folder] does not have [expected worksheets], but [these worksheets].", FxKind)
FxWsChk = MsgLy(M, FfnFn(A), FfnPth(A), CvNy(WsNy0), FxWny(A))
End Function

Function FxWsDt(A, Optional WsNm0$) As Dt
Dim N$: N = FxDftWsNm(A, WsNm0)
Dim Sql$: Sql = FmtQQ("Select * from [?$]", N)
Set FxWsDt = DrsDt(FxSqlDrs(A, Sql), N)
End Function

Function FxWsFny(A, W) As String()
FxWsFny = CatTblFny(CatTbl(FxCat(A), W))
End Function

Function FxwShtTyAy(A, W) As String()
FxwShtTyAy = CatTblShtTyAy(CatTbl(FxCat(A), W))
End Function

Function FxWsIntAy(A, W, Optional F = 0) As Integer()
FxWsIntAy = ARsIntAy(FxWsARs(A, W, F))
End Function

Function FxWsSy(A, W, Optional F = 0) As String()
FxWsSy = ARsSy(FxWsARs(A, W, F))
End Function

Private Sub ZZ_FxFstWsNm()
Debug.Print FxFstWsNm(SampleFx_KE24)
End Sub

Private Sub ZZ_FxWny()
Const Fx$ = "Users\user\Desktop\Invoices 2018-02.xlsx"
D FxWny(Fx)
End Sub

Private Sub ZZ_FxWsNy()
AyDmp FxWsNy(SampleFx_KE24)
End Sub
