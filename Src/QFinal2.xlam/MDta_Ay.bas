Attribute VB_Name = "MDta_Ay"
Option Explicit
Function AyDrs(A) As Drs
Set AyDrs = Drs("Itm", AyDry(A))
End Function

Function AyDry(A) As Variant()
Dim O(), J&
Dim U&: U = UB(A)
ReSz O, U
For J = 0 To U
    O(J) = Array(A(J))
Next
AyDry = O
End Function

Function AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim O(), J&
For J = 0 To UB(A)
    Push O, Array(A(J))
Next
Set AyDt = Dt(DtNm, FldNm, O)
End Function


Function AyGpCntDry(A) As Variant()
If Sz(A) = 0 Then Exit Function
Dim Dup, O(), X, T&, Cnt&
Dup = AyWhDist(A)
For Each X In AyNz(Dup)
    Cnt = AyItmCnt(A, X)
    Push O, Array(X, AyItmCnt(A, X))
    T = T + Cnt
Next
Push O, Array("~Tot", T)
AyGpCntDry = O
End Function

Function AyGpCntDryWhDup(A) As Variant()
AyGpCntDryWhDup = DryWhColGt(AyGpCntDry(A), 1, 1)
End Function

Function AyGpCntFmt(A) As String()
AyGpCntFmt = DryFmtss(AyGpCntDry(A))
End Function


Sub ZZ_AyGpCntFmt()
Dim Ay()
Brw AyGpCntFmt(Ay)
End Sub

Private Sub ZZ_AyCntDry()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = AyCntDry(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
Stop
'AssEqDry Act, Exp
End Sub
