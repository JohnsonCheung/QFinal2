Attribute VB_Name = "MApp_Wrk"
Option Explicit
Private X_W As Database
Function W() As Database
If IsNothing(X_W) Then WEns: WOpn
Set W = X_W
End Function

Function WAcs() As Access.Application
Set WAcs = ApnAcs(Apn)
End Function

Sub WClr()
Dim T, Tny$()
Tny = WTny: If Sz(Tny) = 0 Then Exit Sub
For Each T In Tny
    WDrp T
Next
End Sub

Function WtPrpLoFmlVbl$(T$)
'WtPrpLoFmlVbl = FbtPrpLoFmlVbl(WFb, T)
End Function

Sub WCls()
On Error Resume Next
X_W.Close
Set X_W = Nothing
End Sub

Sub WCrt()
FbCrt WFb
End Sub

Function WDaoCn() As DAO.Connection
'Set WDaoCn = FbDaoCn(WFb)
End Function

Function WDb() As Database
Set WDb = X_W
End Function

Sub WDrp(TT)
'DbttDrp W, TT
End Sub

Sub WEns()
If Not WExist Then WCrt
End Sub

Function WExist() As Boolean
WExist = FfnIsExist(WFb)
End Function

Function WFb$()
WFb = ApnWFb(Apn)
End Function

Function WFbOupTblWb() As Workbook
'Set WFbOupTblWb = FbOupTblWb(WFb)
End Function

Sub WImp(T$, LnkColStr$, Optional WhBExpr$)
If FstChr(T) <> ">" Then Stop
'DbtImpMap W, T, LnkColStr, WhBExpr
End Sub

Sub WImpTbl(TT)
'DbttImp W, TT
End Sub

Sub WIni()
TpImp
WCls
FfnDltIfExist WFb
WOpn
End Sub

Sub WKill()
WCls
FfnDltIfExist WFb
End Sub

Sub WOpn()
'FbEns WFb
'Set X_W = FbDb(WFb)
End Sub

Function WPth$()
WPth = ApnWPth(Apn)
End Function

Sub WPthBrw()
PthBrw WPth
End Sub

Sub WReOpn()
WCls
WOpn
End Sub

Function WrkPth$()
WrkPth = CurPjPth & "WorkingDir\"
End Function

Sub WRun(A)
On Error GoTo X
W.Execute A
Exit Sub
X:
Debug.Print Err.Description
Debug.Print A
Debug.Print "?WStru("""")"
On Error Resume Next
'DbCrtQry W, "Query1", A
Stop
End Sub

Function WtChkCol(T$, LnkColStr$) As String()
'WtChkCol = DbtChkCol(W, T, LnkColStr)
End Function

Function WtColChk(T$, ColLnk$()) As String()
WtColChk = DbtColChk(W, T, ColLnk)
End Function

Sub WtfAddExpr(T$, F$, Expr$)
DbtfAddExpr W, T, F, Expr
End Sub

Function WtFny(T$) As String()
WtFny = DbtFny(W, T)
End Function

Sub WtImp(T$, ColLnk$())
DbtImp W, T, ColLnk
End Sub

Function WTny() As String()
WTny = AySrt(DbTny(W))
End Function

Sub WtRen(T$, ToTbl$)
DbtRen W, T, ToTbl
End Sub

Sub WtRenCol(T$, Fm$, NewColNm$)
DbtRenCol W, T, Fm, NewColNm
End Sub

Sub WtReSeq(T$, ReSeqSpec$)
DbtSpecReSeq W, T, ReSeqSpec
End Sub

Sub WttStruDmp(TT)
D WttStru(TT)
End Sub

Function WtStru$(T$)
WtStru = DbtStru(W, T)
End Function

Function WttStru$(TT)
WttStru = DbttStru(W, TT)
End Function

Function WWbCnStr$()
WWbCnStr = FbOleCnStr(WFb)
End Function


Function WtLnkFb(T$, Fb$) As String()
WtLnkFb = DbtLnkFb(W, T, Fb)
End Function

Function WtLnkFx(T$, Fx$, Optional WsNm$ = "Sheet1") As String()
WtLnkFx = DbtLnkFx(W, T, Fx, WsNm)
End Function

Sub WttLnkFb(TT$, Fb$, Optional Fbtt$)
DbttLnkFb W, TT, Fb$, Fbtt
End Sub
