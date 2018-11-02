Attribute VB_Name = "MVb_Ay_Ins"
Option Explicit

Function AyIns2(A, X1, X2, Optional At&)
Dim O
O = AyReSz(A, At, 2)
Asg X1, O(At)
Asg X2, O(At + 1)
AyIns2 = O
End Function

Function AyIns(A, Optional M, Optional At = 0)
If 0 > At Or At > Sz(A) Then Stop
Dim O
O = AyReSz(A, At)
If Not IsMissing(M) Then
    Asg M, O(At)
End If
AyIns = O
End Function
Private Sub Z_AyIns()
Dim A(), M, At&
'--
A = Array(1, 2, 3, 4, 5)
M = "a"
At = 2
Ept = Array(1, 2, "a", 3, 4, 5)
GoSub Tst
'
Exit Sub
Tst:
    Act = AyIns(A, M, At)
    C
    Return
End Sub
Function AyInsAy(A, B, Optional At&)
Dim O, NB&, J&
NB = Sz(B)
O = AyReSz(A, At, NB)
For J = 0 To NB - 1
    Asg B(J), O(At + J)
Next
AyInsAy = O
End Function

Private Function AyReSz(A, At, Optional Cnt = 1)
If Cnt < 1 Then Stop
Dim P1, P3
    P3 = AyMid(A, At)
    P1 = A
    If At = 0 Then
        Erase P1
    Else
        ReDim Preserve P1(At + Cnt - 1)
    End If
AyReSz = AyAddAp(P1, P3)
End Function
Function AyEmpEle(A)
Dim O: O = A: Erase O
ReDim O(0)
AyEmpEle = O(0)
End Function

Private Sub Z_AyReSz()
Dim Ay(), At&, Cnt&
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Ept = Array(1, Empty, Empty, Empty, 2, 3)
Exit Sub
Tst:
    Act = AyReSz(Ay, At, Cnt)
    Ass IsEqAy(Act, Ept)
End Sub
