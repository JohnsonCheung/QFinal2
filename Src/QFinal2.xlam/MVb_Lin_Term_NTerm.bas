Attribute VB_Name = "MVb_Lin_Term_NTerm"
Option Explicit

Function Lin2T(A) As String()
Lin2T = LinNTerm(A, 2)
End Function

Function Lin3T(A) As String()
Lin3T = LinNTerm(A, 3)
End Function

Function LinNTerm(A, N%) As String()
Dim J%, L$
L = A
For J = 1 To N
    PushI LinNTerm, ShfT(L)
Next
End Function

Function LinTT(ByVal A) As String()
Dim T1$, T2$
T1 = ShfTerm(A)
T2 = ShfTerm(A)
LinTT = ApSy(T1, T2)
End Function
