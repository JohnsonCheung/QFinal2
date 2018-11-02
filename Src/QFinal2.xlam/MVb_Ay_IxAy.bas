Attribute VB_Name = "MVb_Ay_IxAy"
Option Explicit
Function U_IntAy(U&) As Integer()
Dim J&
For J = 0 To U
    PushI U_IntAy, J
Next
End Function

Function U_IxAy(U&) As Long()
Dim O&()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = J
    Next
U_IxAy = O
End Function
Function AyIdx&(A, Itm)
AyIdx = AyIdxFm(A, Itm, 0)
End Function

Function AyIdxFm&(A, Itm, Fm&)
Dim O&
For O = Fm To UB(A)
    If A(O) = Itm Then AyIdxFm = O: Exit Function
Next
AyIdxFm = -1
End Function

Function AyIx&(A, M)
Dim J&
For J = 0 To UB(A)
    If A(J) = M Then AyIx = J: Exit Function
Next
AyIx = -1
End Function

Function AyIxAy(A, SubAy, Optional ChkNotFound As Boolean) As Long()
Dim I
For Each I In AyNz(SubAy)
    PushI AyIxAy, AyIx(A, I)
Next
End Function

Sub AyIxAyAsg(A, IxAy&(), ParamArray OAp())
Dim J%
For J = 0 To UB(IxAy)
    OAp(J) = A(IxAy(J))
Next
End Sub

Sub AyIxAyAsgAp(A, IxAy&(), ParamArray OAp())
Dim J&
For J = 0 To UB(IxAy)
    Asg A(IxAy(J)), OAp(J)
Next
End Sub

Function AyIxAyI(A, B) As Integer()
AyIxAyI = AyIxAyInto(A, B, EmpIntAy)
End Function

Function AyIxAyInto(A, B, OIntoAy)
Dim J&, U&, O
O = OIntoAy
Erase O
U = UB(B)
ReDim O(U)
For J = 0 To U
    O(J) = AyIx(A, B(J))
Next
AyIxAyInto = O
End Function

Function AyIxLblLinPair(A) As String()
'It is 2 line first line is 0 ...
'first line is x0 x1 of A$()
Dim U&: U = UB(A)
If U = -1 Then Exit Function
Dim A1$()
Dim A2$()
ReSz A1, U
ReSz A2, U
Dim O$(), J%, L$, W%
For J = 0 To U
    L = Len(A(J))
    W = Max(L, Len(J))
    A1(J) = AlignL(J, W)
    A2(J) = AlignL(A(J), W)
Next
AyIxLblLinPair = Sy(A1, A2)
End Function
