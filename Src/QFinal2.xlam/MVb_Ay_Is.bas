Attribute VB_Name = "MVb_Ay_Is"
Option Explicit
Function AyIsAllEleEq(A) As Boolean
If Sz(A) = 0 Then AyIsAllEleEq = True: Exit Function
Dim J&
For J = 1 To UB(A)
    If A(0) <> A(J) Then Exit Function
Next
AyIsAllEleEq = True
End Function

Function AyIsAllEleHasVal(A) As Boolean
If Sz(A) = 0 Then Exit Function
Dim I
For Each I In A
    If IsEmp(I) Then Exit Function
Next
AyIsAllEleHasVal = True
End Function

Function AyIsAllEq(A) As Boolean
If Sz(A) <= 1 Then AyIsAllEq = True: Exit Function
Dim A0, J&
A0 = A(0)
For J = 2 To UB(A)
    If A0 <> A(0) Then Exit Function
Next
AyIsAllEq = True
End Function

Function AyIsAllStr(A) As Boolean
Dim K
For Each K In AyNz(A)
    If Not IsStr(K) Then Exit Function
Next
AyIsAllStr = True
End Function

Function AyIsEqSz(A, B) As Boolean
AyIsEqSz = Sz(A) = Sz(B)
End Function
