Attribute VB_Name = "MXls_Z_Lo_Fmt"
Option Explicit
Public Const M_Val_IsNonNum$ = "Lx(?) has Val(?) should be a number"
Public Const M_Val_IsNonLng$ = "Lx(?) has Val(?) should be a 'Long' number"
Public Const M_Val_ShouldBet$ = "Lx(?) has Val(?) should be between [?] and [?]"
Public Const M_Fld_IsInValid$ = "Lx(?) Fld(?) is invalid.  Not found in Fny"
Public Const M_Fld_IsDup$ = "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"
Public Const M_Nm_LinHasNoVal$ = "Lx(?) is Nm-Lin, it has no value"
Public Const M_Nm_NoNmLin$ = "Nm-Lin is missing"
Public Const M_Nm_ExcessLin$ = "LX(?) is excess due to Nm-Lin is found above"
Public Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
Public Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
Public Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"

Const M_Fny$ = "LinTy(?) has these Fld(?) in not Fny"
Const M_Bdr_ExcessFld$ = "These Fld(?) in [Bdr ?] already exists in [Bdr ?], they are skipped in setting border"
Const M_Bdr_ExcessLin$ = "These Fld(?) in [Bdr ?] already exists in [Bdr ?], they are skipped in setting border"
Const M_CorVal$ = "In Lin(?)-Color(?), color cannot convert to long"
Const M_Fld_IsAvg_FndInSum$ = "Lin(?)-Fld(?), which is TAvg-Fld, but also found in TSum-Lx(?)"
Const M_Fld_IsCnt_FndInSum$ = "Lin(?)-Fld(?), which is TCnt-Fld, but also found in TSum-Lx(?)"
Const M_Fld_IsCnt_FndInAvg$ = "Lin(?)-Fld(?), which is TCnt-Fld, but also found in TAvg-Lx(?)"
Const M_Bet_Should2Term = "Lin(?)-Fld(?) is Bet-Line.  It should have 2 terms"
Const M_Bet_InvalidTerm = "Lin(?)-Fld(?) is Bet-Line.  It has invalid term(?)"
Const M_Dup$ = "Lin(?)-Fld(?) is duplicated.  The line is skipped"
Dim A_Lo As ListObject, A_Fny$(), A_LoFmtr$()
Dim Align$(), Bdr$(), Tot, Bet$()
Dim Wdt$(), Fmt$(), Lvl$(), Cor$()
Dim Tit$(), Fml$(), Lbl$()

Private Sub AAMain()
Z_LoFmt
End Sub

Sub AA_2()
Z_ErBet
End Sub

Private Function ErAlign() As String()
ErAlign = SyAddAp(ErAlignLin, ErAlignFny)
End Function

Private Function ErAlignFny() As String()
Dim ErFny$()
    Dim AlignFny$()
    AlignFny = AyWhDist(SSAySy(AyRmvTT(Align)))
    ErFny = AyMinus(AlignFny, A_Fny)
End Function

Private Function ErAlignLin() As String()
ErAlignLin = MsgAlignLin(AyWhExlT1Ay(Align, "Left Right Center"))
End Function

Private Function ErBdr1(X$) As String()
'Return FldAy from Bdr & X
Dim FldssAy$(): FldssAy = SSAySy(AyWhRmvT1(Bdr, X))
End Function

Private Function ErBdr() As String()
ErBdr = SyAddAp(ErBdrExcessFld, ErBdrExcessLin, ErBdrDup, ErBdrFld)
End Function

Private Function ErBdrDup() As String()
ErBdrDup = MsgDup(AyDupT1(Bdr), Bdr)
End Function

Private Function ErBdrExcessFld() As String()
Dim LFny$(), RfNy$(), CFny$()
LFny = ErBdr1("Left")
RfNy = ErBdr1("Right")
CFny = ErBdr1("Center")
PushIAy ErBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, LFny), "Center", "Left")
PushIAy ErBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, RfNy), "Center", "Right")
PushIAy ErBdrExcessFld, FmtQQ(M_Dup, AyMinus(LFny, RfNy), "Left", "Right")
End Function

Private Function ErBdrExcessLin() As String()
Dim L
For Each L In AyNz(AyWhExlT1Ay(Bdr, "Left Right Center"))
    PushI ErBdrExcessLin, FmtQQ(M_Bdr_ExcessLin, L)
Next
End Function

Private Function ErBdrFld() As String()
Dim Fny$(): Fny = SyAddAp(ErBdr1("Left"), ErBdr1("Right"), ErBdr1("Center"))
ErBdrFld = MsgFny(Fny, "Bdr")
End Function

Private Function ErBet() As String()
ErBet = SyAddAp(ErBetDup, ErBetFny, ErBetTermCnt)
End Function

Private Function ErBetDup() As String()
ErBetDup = MsgDup(AyDupT1(Bet), Bet)
End Function

Private Function ErBetFny() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return Er of M_Bet_* if any
End Function

Private Function ErBetTermCnt() As String()
Dim L
For Each L In AyNz(Bet)
    If Sz(SslSy(L)) <> 3 Then
        PushI ErBetTermCnt, MsgBetTermCnt(L, 3)
    End If
Next
End Function

Private Function ErCor() As String()
Dim L$()
L = Cor
ErCor = SyAddAp(ErCorDup(L), ErCorFld(L), ErCorVal(L))
Cor = L
End Function

Private Function ErCorDup(IO$()) As String()

End Function

Private Function ErCorFld(IO$()) As String()

End Function

Private Function ErCorVal1$(L)
Dim Cor$
Cor = LinT1(L)
If IsEmpty(CvColr(L)) Then
    ErCorVal1 = FmtQQ(M_CorVal, L, Cor)
End If
If CanCvLng(Cor) Then Exit Function
End Function

Private Function ErCorVal(IO$()) As String()
Dim Msg$(), Er$(), L
For Each L In IO
    PushI Msg, ErCorVal1(L)
Next
IO = AyWhNoEr(IO, Msg, Er)
End Function

Private Function ErFml() As String()
ErFml = SyAddAp(ErFmlDup, ErFmlFny)
End Function

Private Function ErFmlDup() As String()
ErFmlDup = MsgDup(AyDupT1(Fml), Fml)
End Function

Private Function ErFmlFny() As String()
'ErFmlFny = AyMinus(FmlNy(Fml), A_Fny)
End Function

Private Function ErFmt() As String()

End Function

Private Function ErLbl() As String()

End Function

Private Function ErLvl() As String()

End Function

Private Function ErTit() As String()

End Function

Private Function ErTot() As String()
Dim L
For Each L In AyNz(Tot)
'    A = Avg(J)
'    Ix = AyIx(Sum, A)
'    If Ix >= 0 Then
'        Msg = FmtQQ(M_Fld_IsAvg_FndInSum, AvgLxAy(J), Avg(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    End If
Next
End Function

Private Function ErTot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As Er
'Dim O As New Er
'Dim J%, C$, Ix%, Msg$
'For J = 0 To UB(Cnt)
'    C = Cnt(J)
'    Ix = AyIx(Sum, C)
'    If Ix >= 0 Then
'        Msg = FmtQQ(M_Fld_IsCnt_FndInSum, CntLxAy(J), Cnt(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    Else
'        Ix = AyIx(Avg, C)
'        If Ix >= 0 Then
'            Msg = FmtQQ(M_Fld_IsCnt_FndInAvg, CntLxAy(J), Cnt(J), AvgLxAy(Ix))
'            O.PushMsg Msg
'        End If
'    End If
'Next
'Set ErTot_1 = O
End Function

Private Function ErWdt() As String()
End Function

Private Function HasTot() As Boolean
Dim Lc As ListColumn
For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_HasTot(Lc, FmtSpecLy) Then LoFmtSpecLy_HasTot = True: Exit Function
Next
End Function

Function LoFmt(Lo As ListObject, LoFmtr$()) As ListObject
Set A_Lo = Lo
A_Fny = LoFny(Lo)
A_LoFmtr = LoFmtr

Bdr = X("Bdr")
Align = X("Align")
Tot = X("Tot")

Wdt = X("Wdt")
Fmt = X("Fmt")
Lvl = X("Lvl")
Cor = X("Cor")

Bet = X("Bet")
Fml = X("Fml")
Lbl = X("Lbl")
Tit = X("Tit")

ErRaise CvSy(AyAddAp( _
    ErAlign, ErBdr, ErTot, _
    ErWdt, ErFmt, ErLvl, ErCor, _
    ErFml, ErLbl, ErTit, ErBet))

SetAlign
SetBdr
SetTot

SetWdt
SetLvl
SetCor
SetFmt

SetTit
SetLbl
SetFml
SetBet
End Function

Private Function MsgAlignLin(Ly$()) As String()
If Sz(Ly) Then Exit Function

End Function

Private Function MsgBetTermCnt$(L, NTerm%)

End Function

Private Function MsgDup1(N, Ly$()) As String()
Dim L
For Each L In Ly
    If LinT1(L) = N Then PushI MsgDup1, FmtQQ(M_Dup, L, N)
Next
End Function

Private Function MsgDup(DupNy$(), Ly$()) As String()
Dim N
For Each N In AyNz(DupNy)
    PushIAy MsgDup, MsgDup1(N, Ly)
Next
End Function

Private Function MsgFny(Fny$(), LinTy$) As String()
'Return Msg if given-Fny has some field not in A_Fny
Dim ErFny$(): ErFny = AyMinus(Fny, A_Fny)
If Sz(ErFny) = 0 Then Exit Function
PushI MsgFny, FmtQQ(M_Fny, ErFny, LinTy)
End Function

Sub SampleLoFmtrTpBrw()
Brw SampleLoFmtrTp
End Sub

Private Function SetAlign1(Fldss, A As XlHAlign)
Dim F
For Each F In A_Fny
    If StrLikss(F, Fldss) Then LcSetAlign A_Lo, F, A
Next
End Function

Private Sub SetAlign()
SetAlign1 AyFstRmvT1(Align, "Left"), xlHAlignLeft
SetAlign1 AyFstRmvT1(Align, "Right"), xlHAlignRight
SetAlign1 AyFstRmvT1(Align, "Center"), xlHAlignCenter
End Sub

Private Sub SetBdr()
Dim L$(), R$(), C$()
L = SslSy(JnSpc(AyWhRmvT1(Bdr, "Left")))
R = SslSy(JnSpc(AyWhRmvT1(Bdr, "Right")))
C = SslSy(JnSpc(AyWhRmvT1(Bdr, "Center")))
SetBdrLeft L
SetBdrLeft C
SetBdrRight C
SetBdrRight R
End Sub

Private Function SetBdrLeft(FldLikAy$())
Dim F
For Each F In A_Fny
    If StrLikAy(F, FldLikAy) Then LcSetBdrLeft A_Lo, F
Next
End Function

Private Function SetBdrRight(FldLikAy$())
Dim F
For Each F In A_Fny
    If StrLikAy(F, FldLikAy) Then LcSetBdrRight A_Lo, F
Next
End Function

Private Sub SetBet()
Dim L, C$, X$, Y$
For Each L In AyNz(Tot)
    LinAsg2TRst L, C, X, Y
    A_Lo.ListColumns(C).DataBodyRange.Formula = FmtQQ("=Sum([?]:[?])", X, Y)
Next
End Sub

Private Sub SetCor1(A)
Dim Cor1&, Fldss$, F
LinAsgTRst A, Cor1, Fldss
For Each F In A_Fny
    If StrLikss(F, Fldss) Then LcSetCor A_Lo, F, Cor1
Next
End Sub

Private Sub SetCor()
Dim L
For Each L In AyNz(Cor)
    SetCor1 L
Next
End Sub

Private Sub SetFml()
SetFmlBet

Dim C$, L, Fml1$
For Each L In AyNz(Fml)
    LinAsgTRst L, C, Fml1
    LcSetFml A_Lo, C, Fml1
Next
End Sub

Private Sub SetFmlBet()

End Sub

Private Sub SetFmt1(L)
Dim F
Dim Fmt$, Fldss$
For Each F In A_Fny
    LinAsgTRst L, Fmt, Fldss
    If StrLikss(F, Fldss) Then
        LcSetFmt A_Lo, F, Fmt
        Exit Sub
    End If
Next
End Sub

Private Sub SetFmt()
Dim L
For Each L In AyNz(Fmt)
    SetFmt1 L
Next
End Sub

Private Sub SetLbl()

End Sub

Private Sub SetLvl1(L)
Dim F
Dim Lvl As Byte, Fldss$
For Each F In A_Fny
    LinAsgTRst L, Lvl, Fldss
    If StrLikss(F, Fldss) Then
        LcSetLvl A_Lo, F, Lvl
        Exit Sub
    End If
Next
End Sub

Private Sub SetLvl()
Dim L
For Each L In AyNz(Lvl)
    SetLvl1 L
Next
End Sub

Private Sub SetTit()
LoTitLySet A_Lo, Tit
End Sub

Private Sub SetTot1(FldLikss$, B As XlTotalsCalculation)
Dim F
For Each F In A_Fny
    If StrLikss(F, FldLikss) Then LcSetTot A_Lo, F, B
Next
End Sub

Private Sub SetTot()
SetTot1 AyFstRmvT1(Tot, "Sum"), xlTotalsCalculationSum
SetTot1 AyFstRmvT1(Tot, "Cnt"), xlTotalsCalculationCount
SetTot1 AyFstRmvT1(Tot, "Avg"), xlTotalsCalculationAverage
End Sub

Private Sub SetWdt1(L)
Dim F, W%, Likss$
LinAsgTRst L, W, Likss
For Each F In A_Fny
    If StrLikss(F, Likss) Then LcSetWdt A_Lo, F, W: Exit For
Next
End Sub

Private Sub SetWdt()
Dim L, F, W%, Likss1$
For Each L In AyNz(Wdt)
    SetWdt1 L
Next
End Sub

Private Function X(T1$) As String()
X = AyWhRmvT1(A_LoFmtr, T1)
End Function

Sub Z()
Z_ErBet
Z_LoFmt
Z_SetBdr
End Sub

Private Sub Z_ErBet()
'---------------
A_Fny = SslSy("A B")
Erase Bet
    PushI Bet, "A B C"
    PushI Bet, "A B C"
Ept = EmpSy
    PushIAy Ept, MsgDup(ApSy("A"), Bet)
GoSub Tst
Exit Sub
'---------------
Tst:
    Act = ErBet
    C
    Return
End Sub

Private Sub Z_LoFmt()
Dim Lo As ListObject, LoFmtr$()
'------------
Set Lo = SampleLo
LoFmtr = SampleLoFmtr
GoSub Tst
Exit Sub
Tst:
    LoFmt Lo, LoFmtr
    Return
End Sub

Private Sub Z_SetBdr()
'--
Set A_Lo = SampleLo
'--
Erase Bdr
PushI Bdr, "Left A B C"
PushI Bdr, "Left D E F"
PushI Bdr, "Right A B C"
PushI Bdr, "Center A B C"
GoSub Tst
Tst:
    
    SetBdr      '<==
    Return
End Sub

