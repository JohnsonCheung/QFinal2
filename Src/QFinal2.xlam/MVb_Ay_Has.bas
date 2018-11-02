Attribute VB_Name = "MVb_Ay_Has"
Option Explicit
Function AyHas(A, M) As Boolean
Dim I
For Each I In AyNz(A)
    If I = M Then AyHas = True: Exit Function
Next
End Function

Function AyHasAy(A, Ay) As Boolean
Dim I
For Each I In Ay
    If Not AyHas(A, I) Then Exit Function
Next
AyHasAy = True
End Function

Function AyApHasEle(ParamArray AyAp()) As Boolean
Dim Av(): Av = AyAp
Dim Ay
For Each Ay In AyNz(Av)
    If Sz(Ay) > 0 Then AyApHasEle = True: Exit Function
Next
End Function


Function AyHasAyChk(A, B) As String()
Dim C
C = AyMinus(B, A)
If Sz(C) = 0 Then Exit Function
Er "AyHasAyChk", "[Some-Ele] in [Ay-B] not [Ay-A]", C, B, A
End Function

Function AyHasAyInSeq(A, B) As Boolean
Dim BItm, Ix&
If Sz(B) = 0 Then Stop
For Each BItm In B
    Ix = AyIdxFm(A, BItm, Ix)
    If Ix = -1 Then Exit Function
    Ix = Ix + 1
Next
AyHasAyInSeq = True
End Function

Function AyHasDupEle(A) As Boolean
If Sz(A) = 0 Then Exit Function
Dim Pool: Pool = A: Erase Pool
Dim I
For Each I In A
    If AyHas(Pool, I) Then AyHasDupEle = True: Exit Function
    Push Pool, I
Next
End Function

Function AyHasNegOne(A) As Boolean
Dim V
If Sz(A) = 0 Then Exit Function
For Each V In A
    If V = -1 Then AyHasNegOne = True: Exit Function
Next
End Function

Function AyHasPredPXTrue(A, PX$, P) As Boolean
Dim X
For Each X In AyNz(A)
    If Run(PX, P, X) Then AyHasPredPXTrue = True: Exit Function
Next
End Function

Function AyHasPredXPTrue(A, XP$, P) As Boolean
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In AyNz(A)
    If Run(XP, X, P) Then
        AyHasPredXPTrue = True
        Exit Function
    End If
Next
End Function

Function AyHasSubAy(A, SubAy) As Boolean
Const CSub$ = "AyHasSubAy"
If Sz(A) = 0 Then Exit Function
If Sz(SubAy) = 0 Then Er CSub, "{SubAy} is empty", SubAy
Dim I
For Each I In SubAy
    If Not AyHas(A, I) Then Exit Function
Next
End Function

Sub ZZ_AyHasAyInSeq()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert AyHasAyInSeq(A, B) = True

End Sub

Private Sub ZZ_AyHasDupEle()
Ass AyHasDupEle(Array(1, 2, 3, 4)) = False
Ass AyHasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub
