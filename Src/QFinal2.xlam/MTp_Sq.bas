Attribute VB_Name = "MTp_Sq"
Option Explicit
Public Const SqTpBlkTyss$ = "ER PM SW SQ RM"
Public Const Msg_Sq_1_NotInEDic = "These items not found in ExprDic [?]"
Public Const Msg_Sq_1_MustBe1or0 = "For %?xxx, 2nd term must be 1 or 0"
Sub Z()
Z_SqTpSqy
End Sub

Private Sub AAMain()
Z_SqTpSqy
End Sub

Function ChkErBlk(A() As Blk) As String()

End Function

Function ChkExcessPmBlk(A() As Blk) As String()

End Function

Function ChkExcessSwBlk(A() As Blk) As String()
Dim I, B() As Blk
For Each I In AyNz(A)
    If CvBlk(I).BlkTyStr = "SW" Then
    End If
    
    
Next
End Function

Private Function FndBlkAy1(GpAy() As Gp) As Blk()
Dim I
For Each I In AyNz(GpAy)
    PushObj FndBlkAy1, FndBlk(CvGp(I))
Next
End Function

Function GpAyRmvRmk(A() As Gp) As Gp()
'Dim J%, O() As Gp, M As Gp
'For J = 0 To UB(A)
'    M = GpRmvRmk(A(J))
'    If Sz(M.LnxAy) > 0 Then
'        PushObj O, M
'    End If
'Next
'GpAyRmvRmk = O
End Function

Function FndBlk(A As Gp) As Blk
Set FndBlk = New Blk
With FndBlk
    .BlkTyStr = FndBlkTyStr(A)
    Set .Gp = A
End With
End Function

Function FndBlkTyStr$(A As Gp)
Dim Ly$(): Ly = GpLy(A)
Dim O$
Select Case True
Case LyIsPm(Ly): O = "PM"
Case LyIsSw(Ly): O = "SW"
Case LyIsRm(Ly): O = "RM"
Case LyIsSq(Ly): O = "SQ"
Case Else: O = "ER"
End Select
FndBlkTyStr = O
End Function

Function LnxAy_BlkTyStr$(A() As Lnx)
Dim Ly$(): ' Ly = LnxAy_ZIs(Ly)
Dim O$
Select Case True
Case LyIsPm(Ly): O = "PM"
Case LyIsSw(Ly): O = "SW"
Case LyIsRm(Ly): O = "RM"
Case LyIsSq(Ly): O = "SQ"
Case Else: O = "ER"
End Select
LnxAy_BlkTyStr = O
End Function

Function LnxAy_ErIx_OfDupKey(A() As Lnx) As Integer()

End Function

Function LnxAy_ErIx_OfPfxPercent(A() As Lnx) As Integer()

End Function

Function LnxAy_WhByExclErIxAy(A() As Lnx, ErIxAy%()) As String()
Dim O$(), J%
For J = 0 To UB(A)
    If Not AyHas(ErIxAy, J) Then
        Push O, A(J).Lin
    End If
Next
LnxAy_WhByExclErIxAy = O
End Function

Function LyBlkTyStr$(A$())
Dim O$
Select Case True
Case LyIsPm(A): O = "PM"
Case LyIsSw(A): O = "SW"
Case LyIsRm(A): O = "RM"
Case LyIsSq(A): O = "SQ"
Case Else: O = "ER"
End Select
LyBlkTyStr = O
End Function

Function LyGpAy(Ly$()) As Gp()
Dim O() As Gp, J&, LnxAy() As Lnx, M As Lnx
For J = 0 To UB(Ly)
    Dim Lin$
    Lin = Ly(J)
    If HasPfx(Lin, "==") Then
        If Sz(LnxAy) > 0 Then
            PushObj O, Gp(LnxAy)
        End If
        Erase LnxAy
    Else
        PushObj LnxAy, Lnx(J, Lin)
    End If
Next
If Sz(LnxAy) > 0 Then
    PushObj O, Gp(LnxAy)
End If
LyGpAy = O
End Function


Private Function LyIsPm(A$()) As Boolean
LyIsPm = LyHasMajPfx(A, "%")
End Function

Private Function LyIsRm(A$()) As Boolean
LyIsRm = Sz(A) = 0
End Function

Private Function LyIsSq(A$()) As Boolean
If Sz(A) <> 0 Then Exit Function
Dim L$: L = A(0)
Dim Sy$(): Sy = SslSy("?SEL SEL ?SELDIS SELDIS UPD DRP")
If HasPfxAy(L, Sy) Then LyIsSq = True: Exit Function
End Function

Private Function LyIsSw(A$()) As Boolean
LyIsSw = LyHasMajPfx(A, "?")
End Function

Private Function Rslt_1() As String()
'Return a split-of-SwLnxAy-and-ErLy as SwLnxAyErLy
'by if B_Ay(..).ErLy has Er
'       then put into ErLy    (E$())
'       else put into SwLnxAy (O() As SwLnx)
'Dim E$(), O() As SwLnx
'Dim J%, Er$()
'For J = 0 To U
'    Er = B_Ay(J).ErLy
'    If AyIsEmp(Er) Then
'        PushObj O, B_Ay(J)
'    Else
'        PushAy E, Er
'    End If
'Next
'With Rslt_1
'    .ErLy = E
'    .SwLnxAy = O
'End With
End Function

Function BlkAyWhTySelGp(A() As Blk, BlkTyStr$) As Gp()
Dim J%
For J = 0 To UB(A)
    With A(J)
        If .BlkTyStr = BlkTyStr Then
            PushObj BlkAyWhTySelGp, A(J).Gp
        End If
    End With
Next
End Function

Private Function FndLnxAy(A() As Blk, BlkTyStr$) As Lnx()
Dim J%
For J = 0 To UB(A)
    If A(J).BlkTyStr = BlkTyStr Then FndLnxAy = A(J).Gp.LnxAy: Exit Function
Next
End Function

Private Function FndBlkAy(SqTp$) As Blk()
Dim Ly$():            Ly = SplitCrLf(SqTp)
Dim G() As Gp:         G = LyGpAy(Ly)
Dim G1() As Gp:       G1 = GpAyRmvRmk(G)
FndBlkAy = FndBlkAy1(G1)
End Function

'=======================================
Function SqTpEr(SqTp$) As String()

End Function
Private Function SqTpEr1()

End Function
Function SqTpSqy(SqTp$, OEr$()) As String()
Dim B() As Blk: B = FndBlkAy(SqTp)
Dim PmEr$(), Pm As Dictionary
Dim SwEr$(), StmtSw As Dictionary, FldSw As Dictionary
Dim SqEr$()

Set Pm = FndPm(FndLnxAy(B, "PM"), PmEr)

FndSwAsg FndLnxAy(B, "SW"), Pm, _
    StmtSw, FldSw, SwEr
    
SqTpSqy = FndSqy(BlkAyWhTySelGp(B, "SQ"), Pm, StmtSw, FldSw, SqEr)
OEr = AyAddAp( _
    ChkErBlk(B), _
    ChkExcessSwBlk(B), _
    ChkExcessPmBlk(B), _
    PmEr, SwEr, SqEr)
End Function

Private Function ZIsSw(Ly$()) As Boolean
ZIsSw = LyHasMajPfx(Ly, "?")
End Function

Private Function ZZSqTp$()
Static X$
'If X = "" Then X = MdResStr(Md("W01SqTp"), "SqTp")
ZZSqTp = X
End Function

Private Function ZZSqTpLy() As String()
ZZSqTpLy = SplitCrLf(ZZSqTp)
End Function

Private Sub Z_SqTpSqy()
Dim ActEr$(), SqTp$, EptEr$()
'--
SqTp = SampleSqTp
Ept = ""
EptEr = ApSy("")
GoSub Tst
Exit Sub
Tst:
    Act = SqTpSqy(SqTp, ActEr)
    C
    Ass IsEqAy(ActEr, EptEr)
    Return
End Sub

