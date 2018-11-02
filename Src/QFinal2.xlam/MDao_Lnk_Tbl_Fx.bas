Attribute VB_Name = "MDao_Lnk_Tbl_Fx"
Option Explicit
Sub DbLnkFx(A As Database, Fx$, Optional WsNy0)
Dim W
For Each W In DftWsNy(WsNy0, Fx)
   DbtLnkFx A, W, Fx, W
Next
End Sub

Function FxWsChk(A$, Optional WsNy0 = "Sheet1", Optional FxKind$ = "Excel file") As String()
If FfnNotExist(A) Then FxWsChk = FfnNotExistMsg(A, FxKind): Exit Function
If FxHasWs(A, WsNy0) Then Exit Function
Dim M$
M = FmtQQ("[?] in [folder] does not have [expected worksheets], but [these worksheets].", FxKind)
FxWsChk = MsgLy(M, FfnFn(A), FfnPth(A), CvNy(WsNy0), FxWny(A))
End Function

Function DbtLnkFx(A As Database, T, Fx$, Optional WsNm = "Sheet1") As String()
Dim O$(): O = FxWsChk(Fx, WsNm)
If Sz(O) > 0 Then DbtLnkFx = O: Exit Function
Dim Cn$: Cn = FxDaoCnStr(Fx)
Dim Src$: Src = WsNm & "$"
DbtLnkFx = DbtLnk(A, T, Src, Cn)
End Function
