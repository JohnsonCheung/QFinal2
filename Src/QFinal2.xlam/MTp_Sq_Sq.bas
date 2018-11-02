Attribute VB_Name = "MTp_Sq_Sq"
Option Explicit
Option Compare Text
Private Enum eStmtTy
    eUpdStmt = 1
    eDrpStmt = 2
    eSelStmt = 3
End Enum
Const U_Into$ = "INTO"
Const U_Sel$ = "SEL"
Const U_SelDis$ = "SELECT DISTINCT"
Const U_Fm$ = "FM"
Const U_Gp$ = "GP"
Const U_Wh$ = "WH"
Const U_And$ = "AND"
Const U_Jn$ = "JN"
Const U_LeftJn$ = "LEFT JOIN"
Private Pm As Dictionary
Private StmtSw As Dictionary
Private FldSw As Dictionary
Private X As New Sql_Shared
Sub Z()
Z_EvlSelStmt
Z_FndExpr
Z_FndSqy
Z_FndTblNm
Z_XSel
End Sub
Sub AA()
Z_EvlSelStmt
End Sub
Private Sub AAMain()
Z_FndSqy
End Sub

Private Function EvlDrpStmt$(A$())
End Function

Private Function EvlSelStmt$(A$(), E As Dictionary)
Dim O$()
    Dim I, J%, B$(), L$
    B = AyReverseI(A)
    PushI O, XSel(Pop(B), E)
'    PushI O, X.Into(RmvT1(Pop(B)))
'    PushI O, X.Fm(RmvT1(Pop(B)))
    PushIAy O, XJnOrLeftJn(PopJnOrLeftJn(B), E)
    L = PopWh(B)
    If L <> "" Then
        PushI O, XWh(L, E)
        PushIAy O, XAnd(PopAnd(B), E)
    End If
    PushI O, XGp(PopGp(B), E)
EvlSelStmt = JnCrLf(O)
End Function

Private Function Evl_Stmt$(A$())
Dim Ty As eStmtTy
    Ty = FndStmtTy(A)
If FndIsSkip(A, Ty) Then Exit Function
Dim A1$(), Expr As Dictionary
Set Expr = FndExpr(A, A1)
Dim O$
    Select Case Ty
    Case eUpdStmt: O = EvlUpdStmt(A1, Expr)
    Case eDrpStmt: O = EvlDrpStmt(A1)
    Case eSelStmt: O = EvlSelStmt(A1, Expr)
    Case Else: Stop
    End Select
Evl_Stmt = JnCrLf(O)
End Function

Private Function EvlUpdStmt$(A$(), E As Dictionary)

End Function

Private Function FndActiveFny(A$()) As String()
Dim F
For Each F In A
    If FldSw.Exists(F) Then PushI FndActiveFny, F
Next
End Function

Private Function FndExpr(Ly$(), OLy$()) As Dictionary
Dim Expr$()
AyBrkByEle Ly, "$", OLy, Expr
Set FndExpr = LyDic(Expr)
End Function

Private Function FndExprAy(Fny$(), E As Dictionary) As String()
Dim F, M$
For Each F In Fny
    If E.Exists(F) Then
        M = E(F)
    Else
        M = F
    End If
    PushI FndExprAy, M
Next
End Function

Private Function FndIsSkip(Ly$(), Ty As eStmtTy) As Boolean
FndIsSkip = StmtSw.Exists(FndTblNm(Ly, Ty))
End Function

Private Function FndStmtTy(Ly$()) As eStmtTy
Dim L$
L = UCase(RmvPfx(LinT1(Ly(0)), "?"))
Select Case L
Case "SEL": FndStmtTy = eSelStmt
Case "UPD": FndStmtTy = eUpdStmt
Case "DRP": FndStmtTy = eDrpStmt
Case Else: Stop
End Select
End Function

Private Function FndTblNm$(Ly$(), Ty As eStmtTy)
Select Case Ty
Case eStmtTy.eSelStmt: FndTblNm = FndTblNmSel(Ly)
Case eStmtTy.eUpdStmt: FndTblNm = FndTblNmUpd(Ly(0))
Case Else: Stop
End Select
End Function

Private Function FndTblNmSel$(Ly$())
FndTblNmSel = AyFstRmvT1(Ly, "FM")
End Function

Private Function FndTblNmUpd$(ByVal Lin$)
If RmvPfx(ShfTerm(Lin), "?") <> "upd" Then Stop
FndTblNmUpd = Lin
End Function

Private Function FndValAy(K, E As Dictionary, OValAy$(), OQ$)
'Return true if not found
End Function

Private Function FndValPair(K, E As Dictionary, OV1, OV2)
'Return true if not found
End Function

Private Function IsXXX(A$(), XXX$) As Boolean
IsXXX = UCase(LinT1(A(UB(A)))) = XXX
End Function


Private Function MsgLinTyEr(A As Lnx) As String()


End Function

Private Function MsgMustBeIntoLin$(A As Lnx)

End Function

Private Function MsgMustBeSelorSelDis$(A As Lnx)

End Function

Private Function MsgMustNotHasSpcInTblNmOfIntoLin$(A As Lnx)

End Function


Private Function SampleExprDic() As Dictionary
Dim O$()
PushI O, "A XX"
PushI O, "B BB"
PushI O, "C DD"
PushI O, "E FF"
Set SampleExprDic = LyDic(O)
End Function

Private Function SampleSqLnxAy() As Lnx()
Dim O$()
PushI O, "sel ?MbrCnt RecCnt TxCnt Qty Amt"
PushI O, "into #Cnt"
PushI O, "fm   #Tx"
PushI O, "wh   RecCnt bet @XX @XX"
PushI O, "and  RecCnt bet @XX @XX"

PushI O, "$"
PushI O, "?MbrCnt ?Count(Distinct Mbr)"
PushI O, "RecCnt  Count(*)"
PushI O, "TxCnt   Sum(TxCnt)"
PushI O, "Qty     Sum(Qty)"
PushI O, "Amt     Sum(Amt)"
SampleSqLnxAy = LyLnxAy(O)
End Function
Private Function FndEr(A() As Gp, OLyAy()) As String()

End Function

Function FndSqy(SqGpAy() As Gp, PmDic As Dictionary, StmtSwDic As Dictionary, FldSwDic As Dictionary, OEr$()) As String()
Set Pm = Pm
Set StmtSw = StmtSwDic
Set FldSw = FldSwDic
Dim LyAy()
    OEr = FndEr(SqGpAy, LyAy)
Dim Ly
For Each Ly In AyNz(LyAy)
    PushI FndSqy, Evl_Stmt(CvSy(Ly))
Next
End Function
Function MsgAndLinOp_ShouldBe_BetOrIn(A)

End Function
Private Function XAnd(A$(), E As Dictionary)
'and f bet xx xx
'and f in xx
Dim F$, I, L$, Ix%, M As Lnx
For Each I In AyNz(A)
    Set M = I
    LnxAsg M, L, Ix
    If ShfTerm(L) <> "and" Then Stop
    F = ShfTerm(L)
    Select Case ShfTerm(L)
    Case "bet":
    Case "in"
    Case Else: Stop
    End Select
Next
End Function


Private Function XGp$(L$, E As Dictionary)
If L = "" Then Exit Function
Dim ExprAy$(), Ay$()
Stop
'    ExprAy = DicSelIntoSy(EDic, Ay)
'XGp = SqpGp(ExprAy)
End Function

Private Function XJnOrLeftJn(A$(), E As Dictionary) As String()

End Function

Private Function PopJnOrLeftJn(A$()) As String()
PopJnOrLeftJn = PopMulXorYOpt(A, U_Jn, U_LeftJn)
End Function

Private Function PopXXXOpt$(A$(), XXX$)
'Return No-Q-T1-of-LasEle-of-A$() if No-Q-T1-of it = XXX else return ''
If Sz(A) = 0 Then Exit Function
PopXXXOpt = PopXXX(A, XXX)
End Function

Private Function PopXXX$(A$(), XXX$)
'Return No-Q-T1-of-LasEle-of-A$() if No-Q-T1-of it = XXX else Stop
Dim L$: L = A(UB(A))
If RmvPfx(LinT1(L), "?") = XXX Then
    PopXXX = RmvT1(L)
    Pop A
End If
End Function

Private Function PopGp$(A$())
PopGp = PopXXXOpt(A, U_Gp)
End Function

Private Function PopWh$(A$())
PopWh = PopXXXOpt(A, U_Wh)
End Function

Private Function PopAnd(A$()) As String()
PopAnd = PopMulXXX(A, U_And)
End Function

Private Function PopXorYOpt$(A$(), X$, Y$)
Dim L$
L = PopXXXOpt(A, X): If L <> "" Then PopXorYOpt = L: Exit Function
PopXorYOpt = PopXXXOpt(A, Y)
End Function

Private Function PopMulXorYOpt(A$(), X$, Y$) As String()
Dim J%, L$
While Sz(A) > 0
    J = J + 1: If J > 1000 Then Stop
    L = PopXorYOpt(A, X, Y)
    If L = "" Then Exit Function
    PushI PopMulXorYOpt, L
Wend
End Function

Private Function PopMulXXX(A$(), XXX$) As String()
Dim J%
While Sz(A) > 0
    J = J + 1: If J > 1000 Then Stop
    If Not IsXXX(A, XXX) Then Exit Function
    PushObj PopMulXXX, Pop(A)
Wend
End Function

Private Function XSel$(A$, E As Dictionary)
Dim Fny$()
    Dim T1$, L$
    L = A
    T1 = RmvPfx(ShfTerm(L), "?")
    Fny = XSelFny(SslSy(L), FldSw)
Select Case T1
'Case U_Sel:    XSel = X.Sel_Fny_EDic(Fny, E)
'Case U_SelDis: XSel = X.Sel_Fny_EDic(Fny, E, IsDis:=True)
Case Else: Stop
End Select
End Function
Private Function XSelFny(Fny$(), FldSw As Dictionary) As String()
Dim F
For Each F In Fny
    If FstChr(F) = "?" Then
        If Not FldSw.Exists(F) Then Stop
        If FldSw(F) Then PushI XSelFny, F
    Else
        PushI XSelFny, F
    End If
Next
End Function

Private Function XSet(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XUpd(A As Lnx, E As Dictionary, OEr$())

End Function
Private Function XWh$(L$, E As Dictionary)
'L is following
'  ?Fld in @ValLis  -
'  ?Fld bet @V1 @V2
Dim F$, ValAy$(), V1, V2, IsBet As Boolean
If IsBet Then
    If Not FndValPair(F, E, V1, V2) Then Exit Function
    'XWh = X.WhBet(F, V1, V2)
    Exit Function
End If
'If Not FndValAy(F, E, ValAy, Q) Then Exit Function
'XWh = X.WhFldInAy(F, ValAy)
End Function

Private Function XWhBetNbr$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XWhExpr(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function XWhInNbrLis$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Sub Z_EvlSelStmt()
Dim E As Dictionary, Ly$()
'---
Erase Ly
    Push Ly, "?XX Fld-XX"
    Push Ly, "BB Fld-BB-LINE-1"
    Push Ly, "BB Fld-BB-LINE-2"
    Set E = LyDic(Ly)           '<== Set ExprDic
Erase Ly
    Set FldSw = New Dictionary
    FldSw.Add "?XX", False       '<=== Set FldSw
Erase Ly
    Erase Ly
    PushI Ly, "sel ?XX BB CC"
    PushI Ly, "into #AA"
    PushI Ly, "fm   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "wh   A bet $a $b"
    PushI Ly, "and  B in $c"
    PushI Ly, "gp   D C"        '<== SqLy
GoSub Tst
Exit Sub
Tst:
    Act = EvlSelStmt(Ly, E)
    C
    Return
End Sub

Private Sub Z_FndExpr()
Dim Ly$(), ActLy1$(), EptLy1$()
Dim D As New Dictionary
'-----

Erase Ly
PushI Ly, "aaa bbb"
PushI Ly, "111 222"
PushI Ly, "$"
PushI Ly, "A B0"
PushI Ly, "A B1"
PushI Ly, "A B2"
PushI Ly, "B B0"
Erase EptLy1
PushI EptLy1, "aaa bbb"
PushI EptLy1, "111 222"
D.RemoveAll
    D.Add "A", JnCrLf(SslSy("B0 B1 B2"))
    D.Add "B", "B0"
    Set Ept = D
GoSub Tst
Exit Sub
Tst:
    Set Act = FndExpr(Ly, ActLy1)
    Ass IsEqDic(CvDic(Act), CvDic(Ept))
    Ass IsEqAy(ActLy1, EptLy1)
    Return
End Sub

Private Sub Z_FndSqy()
Dim A() As Gp, Pm As Dictionary, StmtSw As Dictionary, FldSw As Dictionary, Er$()
GoSub Dta1
GoSub Tst
Return
Tst:
    Act = FndSqy(A, Pm, StmtSw, FldSw, Er)
    C
    Return
Dta1:
    Return
End Sub

Private Sub Z_XSel()
Dim A$, E As Dictionary
A = "dsklfj"
Set E = SampleExprDic
GoSub Tst
Exit Sub
Tst:
    Act = XSel(A, E)
    C
    Return
End Sub

Private Sub Z_FndTblNm()
Dim Ly$(), Ty As eStmtTy
'---
PushI Ly, "sel sdflk"
PushI Ly, "fm AA BB"
Ept = "AA BB"
Ty = eSelStmt
GoSub Tst
'---
Erase Ly
PushI Ly, "?upd XX BB"
PushI Ly, "fm dsklf dsfl"
Ept = "XX BB"
Ty = eUpdStmt
GoSub Tst
Exit Sub
Tst:
    Act = FndTblNm(Ly, Ty)
    C
    Return
End Sub
