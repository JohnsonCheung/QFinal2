Attribute VB_Name = "MDta_X_Dry"
Option Explicit
Const CMod$ = "MDta_X_Dry."
Function AyConst_ValConstDry(A, Constant) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), I
For Each I In A
   Push O, Array(I, Constant)
Next
AyConst_ValConstDry = O
End Function

Function DryIsEq(A(), B()) As Boolean
DryIsEq = IsEqAy(A, B)
End Function

Function DryWhColInAy(A(), ColIx%, InAy) As Variant()
Const CSub$ = "DryWhColInAy"
If Not IsArray(InAy) Then Er CSub, "[InAy] is not Array, but [TypeName]", InAy, TypeName(InAy)
If Sz(InAy) = 0 Then DryWhColInAy = A: Exit Function
Dim Dr
For Each Dr In AyNz(A)
    If AyHas(InAy, Dr(ColIx)) Then PushI DryWhColInAy, Dr
Next
End Function
Sub C3DryDo(C3Dry(), ABC$)
If Sz(C3Dry) = 0 Then Exit Sub
Dim Dr
For Each Dr In C3Dry
    Run ABC, Dr(0), Dr(1), Dr(2)
Next
End Sub

Sub C4DryDo(C4Dry(), ABCD$)
If Sz(C4Dry) = 0 Then Exit Sub
Dim Dr
For Each Dr In C4Dry
    Run ABCD, Dr(0), Dr(1), Dr(2), Dr(3)
Next
End Sub

Function DotNyDry(DotNy$()) As Variant()
If Sz(DotNy) = 0 Then Exit Function
Dim O(), I
For Each I In DotNy
   With Brk1(I, ".")
       Push O, ApSy(.S1, .S2)
   End With
Next
DotNyDry = O
End Function
Sub ZZ_DryFmt()
AyDmp DryFmt(SampleDry1)
End Sub


Function DryWhColHasDup(A(), ColIx%) As Variant()
Dim B(): B = AyWhDup(DryCol(A, ColIx))
DryWhColHasDup = DryWhColInAy(A, ColIx, B)
End Function

Function DryFstColEqV(A(), ColIx%, V)
Dim Dr
For Each Dr In AyNz(A)
    If Dr(ColIx) = V Then DryFstColEqV = Dr: Exit Function
Next
End Function

Private Function Dry_MgeIx&(Dry(), Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   Dry_MgeIx = O
   Exit Function
Nxt:
Next
Dry_MgeIx = -1
End Function

Function DryAddCC(A(), C1, C2) As Variant()
DryAddCC = DryAddCol(DryAddCol(A, C1), C2)
End Function

Function DryAddCol(A(), C) As Variant()
Dim UCol%, R&, Dr, O()
O = AyReSz(O, A)
UCol = DryNCol(A)
For Each Dr In AyNz(A)
    ReDim Preserve Dr(UCol)
    Dr(UCol) = C
    O(R) = Dr
    R = R + 1
Next
DryAddCol = O
End Function

Function DryAddConstCol(Dry(), ConstVal) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim N%
   N = Sz(Dry(0))
Dim O()
   Dim Dr, J&
   ReDim O(UB(Dry))
   For Each Dr In Dry
       ReDim Preserve Dr(N)
       Dr(N) = ConstVal
       O(J) = Dr
       J = J + 1
   Next
DryAddConstCol = O
End Function

Function DryAddValIdCntCol(A, ColIx) As Variant() ' Add 2 col at end (Id and Cnt) according to col(ColIx)
Dim O(), NCol%, Dr, R&, D As Dictionary, UCol%, IdCnt&()
O = A
UCol = DryNCol(O) + 1   ' The UCol after add
Set D = DryColSeqCntDic(A, ColIx)
For Each Dr In A
    ReDim Preserve Dr(UCol)
    If Not D.Exists(Dr(ColIx)) Then Stop
    IdCnt = D(ColIx)
    Dr(UCol - 1) = IdCnt(0)
    Dr(UCol) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
DryAddValIdCntCol = O
End Function

Function DryAddValIdCol(A(), ValCol) As Variant()
Dim NCol%, Dic As Dictionary, O(), Dr, IdCnt, R&
NCol = DryNCol(A)
Set Dic = AyDistIdCntDic(DryCol(A, ValCol))
O = AyReSz(O, A)
For Each Dr In AyNz(A)
    ReDim Preserve Dr(NCol + 1)
    IdCnt = Dic(Dr(ValCol))
    Dr(NCol) = IdCnt(0)
    Dr(NCol + 1) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
DryAddValIdCol = O
End Function

Sub DryBrw(A, Optional MaxColWdt% = 100, Optional BrkColIx% = -1)
AyBrw DryFmt(A, MaxColWdt, BrkColIx)
End Sub

Function DryCntDic(A, KeyColIx%) As Dictionary
Dim O As New Dictionary
Dim J%, Dr, K
For J = 0 To UB(A)
    Dr = A(J)
    K = Dr(KeyColIx)
    If O.Exists(K) Then
        O(K) = O(K) + 1
    Else
        O.Add K, 1
    End If
Next
Set DryCntDic = O
End Function

Function DryCol(A, ColIx) As Variant()
DryCol = DryColInto(A, ColIx, Array())
End Function

Function DryColInto(A, ColIx, OInto)
Dim O, J&, Dr, U&
O = AyReSz(OInto, A)
For Each Dr In AyNz(A)
    If UB(Dr) >= ColIx Then
        O(J) = Dr(ColIx)
    End If
    J = J + 1
Next
DryColInto = O
End Function

Function DryColSeqCntDic(A, ColIx) As Dictionary
Set DryColSeqCntDic = AySeqCntDic(DryCol(A, ColIx))
End Function

Function DryColCntDic(A, ColIx) As Dictionary
Set DryColCntDic = AyCntDic(DryCol(A, ColIx))
End Function

Function DryColSqlTy$(A(), ColIx%)
Dim O As VbVarType, Dr, V, T As VbVarType
For Each Dr In A
    If UB(Dr) >= ColIx Then
        V = Dr(ColIx)
        T = VarType(V)
        If T = vbString Then
            If Len(V) > 255 Then DryColSqlTy = "Memo": Exit Function
        End If
        O = MaxVbTy(O, T)
    End If
Next
DryColSqlTy = VbTySqlTy(O)
End Function

Sub DryDmp(A())
AyDmp DryFmtss(A)
End Sub

Sub DryDmp1(A())
DryFmtssDmp A
End Sub
Sub DryFmtssDmp(A())
D DryFmtss(A)
End Sub


Function DryMapJnDot(A()) As String()
Dim Dr
For Each Dr In AyNz(A)
    PushI DryMapJnDot, JnDot(Dr)
Next
End Function
Function DryGpAy(A, Kix%, Gix%) As Variant()
If Sz(A) = 0 Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In A
    K = Dr(Kix)
    Gp = Dr(Gix)
    O_Ix = AyIx(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryGpAy = O
End Function

Function DryGpDic(A, K%, G%) As Dictionary
Dim Dr, U&, O As New Dictionary, KK, GG, Ay()
U = UB(A): If U = -1 Then Exit Function
For Each Dr In A
    KK = Dr(K)
    GG = Dr(G)
    If O.Exists(KK) Then
        Ay = O(KK)
        Push Ay, GG
        O(KK) = Ay
    Else
        O.Add KK, Array(GG)
    End If
Next
Set DryGpDic = O
End Function

Function DryGpFlat(A, K%, G%) As Variant()
DryGpFlat = Aydic_to_KeyCntMulItmColDry(DryGpDic(A, K, G))
End Function

Function DryInsC4(A, C1, C2, C3, C4) As Variant()
Dim Dr, O(), C4Ay()
If Sz(A) = 0 Then Exit Function
C4Ay = Array(C1, C2, C3, C4)
For Each Dr In A
    Push O, AyInsAy(Dr, C4Ay)
Next
DryInsC4 = O
End Function

Function DryInsCC(A, C1, C2) As Variant()
Dim Dr, O(), CCAy()
If Sz(A) = 0 Then Exit Function
CCAy = Array(C1, C2)
For Each Dr In A
    Push O, AyInsAy(Dr, CCAy)
Next
DryInsCC = O
End Function

Function DryInsCCC(A, C1, C2, C3) As Variant()
Dim Dr, O(), C3Ay()
If Sz(A) = 0 Then Exit Function
C3Ay = Array(C1, C2, C3)
For Each Dr In A
    Push O, AyInsAy(Dr, C3Ay)
Next
DryInsCCC = O
End Function

Function DryInsCol(A, C, Optional Ix&) As Variant()
Dim Dr
For Each Dr In A
    PushI DryInsCol, AyIns(Dr, C, At:=Ix)
Next
End Function

Function DryInsConst(A, C, Optional At& = 0) As Variant()
Dim O(), Dr
If Sz(A) = 0 Then Exit Function
For Each Dr In A
    Push O, AyIns(Dr, C, At)
Next
DryInsConst = O
End Function

Function DryIntCol(A, ColIx%) As Integer()
DryIntCol = DryColInto(A, ColIx, EmpIntAy)
End Function

Function DryIsBrkAtDrIx(Dry, DrIx&, BrkColIx%) As Boolean
If Sz(Dry) = 0 Then Exit Function
If DrIx = 0 Then Exit Function
If DrIx = UB(Dry) Then Exit Function
If Dry(DrIx)(BrkColIx) = Dry(DrIx - 1)(BrkColIx) Then Exit Function
DryIsBrkAtDrIx = True
End Function

Function DryKeyGpAy(Dry(), K_Ix%, Gp_Ix%) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In Dry
    K = Dr(K_Ix)
    Gp = Dr(Gp_Ix)
    O_Ix = AyIx(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryKeyGpAy = O
End Function


Function DryMge(Dry, MgeIx%, Sep$) As Variant()
Dim O(), J%
Dim Ix%
For J = 0 To UB(Dry)
   Ix = DryMgeIx(O, Dry(J), MgeIx)
   If Ix = -1 Then
       Push O, Dry(J)
   Else
       O(Ix)(MgeIx) = O(Ix)(MgeIx) & Sep & Dry(J)(MgeIx)
   End If
Next
DryMge = O
End Function

Function DryMgeIx&(Dry, Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   DryMgeIx = O
   Exit Function
Nxt:
Next
DryMgeIx = -1
End Function

Function DryNCol%(A)
Dim O%, Dr
For Each Dr In AyNz(A)
    O = Max(O, Sz(Dr))
Next
DryNCol = O
End Function

Function DryPkMinus(A, B, PkIxAy&()) As Variant()
Dim AK(): AK = DrySel(A, PkIxAy)
Dim BK(): BK = DrySel(B, PkIxAy)
Dim CK(): CK = DryPkMinus(AK, BK, PkIxAy)
DryPkMinus = DryWhIxAyValAy(A, PkIxAy, CK)
End Function

Function DryReOrd(Dry, PartialIxAy&()) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim Dr, O()
For Each Dr In Dry
   Push O, AyReOrd(Dr, PartialIxAy)
Next
DryReOrd = O
End Function

Function DryRmvCol(A, ColIx&) As Variant()
Dim X
For Each X In AyNz(A)
    PushI DryRmvCol, AyRmvEleAt(X, ColIx)
Next
End Function

Function DryRmvColByIxAy(A, IxAy) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), Dr
For Each Dr In A
   Push O, AyWhExlIxAy(Dr, IxAy)
Next
DryRmvColByIxAy = O
End Function

Function DryRowCnt&(Dry, ColIx&, EqVal)
If Sz(Dry) = 0 Then Exit Function
Dim J&, O&, Dr
For Each Dr In Dry
   If Dr(ColIx) = EqVal Then O = O + 1
Next
DryRowCnt = O
End Function

Function DrySq(A) As Variant()
Dim O(), C%, R&, Dr
Dim NC%, NR&
NC = DryNCol(A)
NR = Sz(A)
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
    Dr = A(R - 1)
    For C = 1 To Min(Sz(Dr), NC)
        O(R, C) = Dr(C - 1)
    Next
Next
DrySq = O
End Function

Function DryStrCol(A, Optional ColIx% = 0) As String()
DryStrCol = DryColInto(A, ColIx, EmpSy)
End Function

Function DrySy(A, Optional ColIx% = 0) As String()
DrySy = DryStrCol(A, ColIx)
End Function

Function DryWh(Dry(), ColIx%, EqVal) As Variant()
Dim O()
Dim J&
For J = 0 To UB(Dry)
   If Dry(J)(ColIx) = EqVal Then Push O, Dry(J)
Next
DryWh = O
End Function

Function DryWhCCNe(A, C1, C2) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C1) <> Dr(C2) Then PushI DryWhCCNe, Dr
Next
End Function

Sub AssEqDry(A(), B())
If Not DryIsEq(A, B) Then Stop
End Sub


Function DryWdt(A()) As Integer()
Dim J%
For J = 0 To DryNCol(A) - 1
    Push DryWdt, AyWdt(DryCol(A, J))
Next
End Function

Function DryWdtAy(A(), Optional MaxColWdt% = 100) As Integer()
Const CSub$ = CMod & "DryWdtAy"
If Sz(A) = 0 Then Exit Function
Dim O%()
   Dim Dr, UDr%, U%, V, L%, J%
   U = -1
   For Each Dr In A
       If Not IsSy(Dr) Then Er CSub, "This routine should call ACvFmtEachCell first so that each cell is ValCellStr as a string.|Now some Dr in given-A is not a StrAy, but[" & TypeName(Dr) & "]"
       UDr = UB(Dr)
       If UDr > U Then ReDim Preserve O(UDr): U = UDr
       If Sz(Dr) = 0 Then GoTo Nxt
       For J = 0 To UDr
           V = Dr(J)
           L = Len(V)

           If L > O(J) Then O(J) = L
       Next
Nxt:
   Next
Dim M%
    M = Limit(MaxColWdt, 1, 200)
For J = 0 To UB(O)
   If O(J) > M Then O(J) = M
Next
DryWdtAy = O
End Function

Function DryWhColEq(A, C%, V) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C) = V Then PushI DryWhColEq, Dr
Next
End Function

Function DryWhColGt(A, C%, V) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    If Dr(C) > V Then PushI DryWhColGt, Dr
Next
End Function

Function DryWhDup(A, Optional ColIx% = 0) As Variant()
Dim Dup, Dr, O()
Dup = AyWhDup(DryCol(A, ColIx))
For Each Dr In A
    If AyHas(Dup, Dr(ColIx)) Then Push O, Dr
Next
DryWhDup = O
End Function

Function DryWhIxAyValAy(A, IxAy, ValAy) As Variant()
Dim Dr
For Each Dr In A
    If IsEqAy(DrSel(Dr, IxAy), ValAy) Then PushI DryWhIxAyValAy, Dr
Next
End Function

Function DryDistCol(A() As Variant, ColIx%)
DryDistCol = AyWhDist(DryCol(A, ColIx))
End Function

Function DryDistSy(A() As Variant, ColIx%) As String()
DryDistSy = AyWhDist(DrySy(A, ColIx))
End Function

Function DryJnDotSy(A() As Variant) As String()
Dim Dr
For Each Dr In AyNz(A)
    PushI DryJnDotSy, JnDot(Dr)
Next
End Function
