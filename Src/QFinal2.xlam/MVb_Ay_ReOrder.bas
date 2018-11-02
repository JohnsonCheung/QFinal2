Attribute VB_Name = "MVb_Ay_ReOrder"
Option Explicit
Function AyReOrd(A, PartialIxAy)
Dim Ay, Ix
    Ay = AyCln(A)
    For Each Ix In PartialIxAy
        PushI Ay, A(Ix)
    Next
AyReOrd = AyReOrdAy(A, Ay)
End Function

Function AyReOrdAy(A, SubAy)
If Not AyHasAy(A, SubAy) Then Stop
AyReOrdAy = AyAdd(SubAy, AyMinus(A, SubAy))
End Function
