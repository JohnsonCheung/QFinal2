Attribute VB_Name = "MVb_Lin_Term_NTermRst"
Option Explicit

Function Lin1TRst(A) As String()
Lin1TRst = LinNTermRst(A, 1)
End Function

Function Lin2TRst(A) As String()
Lin2TRst = LinNTermRst(A, 2)
End Function

Function Lin3TRst(A) As String()
Lin3TRst = LinNTermRst(A, 3)
End Function

Function Lin4TRst(A) As String()
Lin4TRst = LinNTermRst(A, 4)
End Function

Function LinNTermRst(A, N%) As String()
Dim L$, J%
L = A
For J = 1 To N
    PushI LinNTermRst, ShfT(L)
Next
PushI LinNTermRst, L
End Function

Private Sub Z_LinNTermRst()
Dim A$
A = "  [ksldfj ]":  Ept = "ksldfj ": GoSub Tst
A = "  [ ksldfj ]": Ept = " ksldf ": GoSub Tst
A = "  [ksldfj]":  Ept = "ksldf": GoSub Tst
Exit Sub
Tst:
    Act = LinT1(A)
    C
    Return
End Sub
