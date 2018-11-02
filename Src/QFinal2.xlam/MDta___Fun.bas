Attribute VB_Name = "MDta___Fun"
Option Explicit

Function ApDtAy(ParamArray Ap()) As Dt()
Dim Av(): Av = Ap
ApDtAy = AyInto(Av, EmpDtAy)
End Function


Function LinSimTyAyDr(A, B() As eSimTy) As Variant()
End Function
Function SqRowDr(A, R) As Variant()
Dim C%
For C = 1 To UBound(A, 2)
    PushI SqRowDr, A(R, C)
Next
End Function

Function SqLy(A) As String()
Dim R%
For R = 1 To UBound(A, 1)
    Push SqLy, JnSpc(SqRowDr(A, R))
Next
End Function

Private Sub ZZ_ItrPrpDrs()
DrsDmp ItrPrpDrs(Application.Vbe.VBProjects, "Name Type")
End Sub

Function SqAlign(Sq(), W%()) As Variant()
If UBound(Sq, 2) <> Sz(W) Then Stop
Dim C%, R%, Wdt%, O
O = Sq
For C = 1 To UBound(Sq, 2) - 1 ' The last column no need to align
    Wdt = W(C - 1)
    For R = 1 To UBound(Sq, 1)
        O(R, C) = AlignL(Sq(R, C), Wdt)
    Next
Next
SqAlign = O
End Function
