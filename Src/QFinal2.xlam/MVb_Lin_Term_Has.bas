Attribute VB_Name = "MVb_Lin_Term_Has"
Option Explicit

Function LinHas2T(A, T1, T2) As Boolean
Dim L$: L = A
If ShfT(L) <> T1 Then Exit Function
If ShfT(L) <> T2 Then Exit Function
LinHas2T = True
End Function

Function LinHas3T(A$, T1, T2, T3) As Boolean
Dim L$: L = A
If ShfT(L) <> T1 Then Exit Function
If ShfT(L) <> T2 Then Exit Function
If ShfT(L) <> T3 Then Exit Function
LinHas3T = True
End Function

Function LinHasT1(A, T1) As Boolean
LinHasT1 = LinT1(A) = T1
End Function

Function LinHasT2(A, T2) As Boolean
LinHasT2 = LinT2(A) = T2
End Function
