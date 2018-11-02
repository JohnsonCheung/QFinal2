Attribute VB_Name = "MVb_Ay_Brk"
Option Explicit
Function AyBrk3ByIx(A, FmIx&, ToIx&)
AyBrk3ByIx = AyFTIxBrk(A, FTIx(FmIx, ToIx))
End Function

Sub AyBrkByEle(A, Ele, OAy1, OAy2)
OAy1 = AyCln(A)
OAy2 = AyCln(A)
Dim J%
For J = 0 To UB(A)
    If A(J) = Ele Then Exit For
    PushI OAy1, A(J)
Next
For J = J + 1 To UB(A)
    PushI OAy2, A(J)
Next
End Sub

Function AyBrkInto3Ay(A, FmIx&, ToIx&) As Variant()
Dim O(2)
O(0) = AyWhFmIxToIx(A, 0, FmIx - 1)
O(1) = AyWhFmIxToIx(A, FmIx, ToIx)
O(2) = AyWhFm(A, ToIx + 1)
AyBrkInto3Ay = O
End Function

Function AyFTIxBrk(A, B As FTIx)
AyFTIxBrk = AyBrkInto3Ay(A, B.FmIx, B.ToIx)
End Function
