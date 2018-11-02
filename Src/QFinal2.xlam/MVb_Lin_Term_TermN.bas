Attribute VB_Name = "MVb_Lin_Term_TermN"
Option Explicit

Function LinT1$(A)
LinT1 = LinTermN(A, 1)
End Function

Function LinT2$(A)
LinT2 = LinTermN(A, 2)
End Function

Function LinT3$(A)
LinT3 = LinTermN(A, 3)
End Function

Function LinTermN$(A, N%)
Dim L$, J%
L = A
For J = 1 To N - 1
    ShfTerm L
Next
LinTermN = ShfTerm(L)
End Function

Sub Z_LinTermN()
Dim N%, A$
N = 1: A = "a b c": Ept = "a": GoSub Tst
N = 2: A = "a b c": Ept = "b": GoSub Tst
N = 3: A = "a b c": Ept = "c": GoSub Tst
Exit Sub
Tst:
    Act = LinTermN(A, N)
    C
    Return
End Sub
