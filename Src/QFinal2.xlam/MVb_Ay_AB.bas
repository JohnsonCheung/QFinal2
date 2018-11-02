Attribute VB_Name = "MVb_Ay_AB"
Option Explicit
Function AyabAdd(A, B, Optional Sep$) As String()
Dim O$(), J&, U&
U = UB(A): If U <> UB(B) Then Stop
If U = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = A(J) & Sep & B(J)
Next
AyabAdd = O
End Function

Function AyabAddWSpc(A, B) As String()
AyabAddWSpc = AyabAdd(A, B, " ")
End Function

Function AyabDic(A, B) As Dictionary
Dim N1&, N2&
N1 = Sz(A)
N2 = Sz(B)
If N1 <> N2 Then Stop
Set AyabDic = New Dictionary
Dim J&, X
For Each X In AyNz(A)
    AyabDic.Add X, B(J)
    J = J + 1
Next
End Function

Function AyabFmt(A, B) As String()
AyabFmt = S1S2AyFmt(AyabS1S2Ay(A, B))
End Function

Function AyabMapInto(A, B, FunAB$, OInto)
Dim J&, U&, O
O = OInto
U = Min(UB(A), UB(B))
If U >= 0 Then ReDim O(U)
For J = 0 To U
    O(J) = Run(FunAB, A(J), B(J))
Next
AyabMapInto = O
End Function

Function AyabMapSy(A, B, FunAB$) As String()
AyabMapSy = AyabMapInto(A, B, FunAB, EmpSy)
End Function

Function AyabNonEmpBLy(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
AyabNonEmpBLy = O
End Function

Sub AyabSetSamMax(A, B)
Dim U1&, U2&
U1 = UB(A)
U2 = UB(B)
Select Case True
Case U1 > U2: ReDim Preserve B(U1)
Case U1 < U2: ReDim Preserve A(U2)
End Select
End Sub

Sub AyabEqChk(A, B, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act")
AssChk AyEqChk(A, B, Ay1Nm, Ay2Nm)
End Sub
