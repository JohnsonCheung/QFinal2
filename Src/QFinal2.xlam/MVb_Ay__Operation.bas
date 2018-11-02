Attribute VB_Name = "MVb_Ay__Operation"
Option Explicit

Function AyIntersect(A, B)
AyIntersect = AyCln(A)
If Sz(A) = 0 Then Exit Function
If Sz(A) = 0 Then Exit Function
Dim V
For Each V In A
    If AyHas(B, V) Then PushI AyIntersect, V
Next
End Function
Function AyMin(A)
Dim O, J&
If Sz(A) = 0 Then Exit Function
O = A(0)
For J = 1 To UB(A)
    If A(J) < O Then O = A(J)
Next
AyMin = O
End Function

Function AyMinus(A, B)
If Sz(B) = 0 Then AyMinus = A: Exit Function
AyMinus = AyCln(A)
If Sz(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not AyHas(B, V) Then
        PushI AyMinus, V
    End If
Next
End Function

Function AyMinusAp(A, ParamArray AyAp())
Dim O: O = A
Dim Av(): Av = AyAp
Dim Ay
For Each Ay In Av
    If Sz(O) = 0 Then GoTo X
    O = AyMinus(O, Ay)
Next
X:
AyMinusAp = O
End Function

Function AyMax(A)
Dim O, I
For Each I In AyNz(A)
    If I > O Then O = I
Next
AyMax = O
End Function

Function AyMaxSz%(A)
If Sz(A) = 0 Then Exit Function
Dim O&, I, S&
For Each I In A
    O = Max(O, Sz(I))
Next
AyMaxSz = O
End Function
