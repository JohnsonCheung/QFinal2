Attribute VB_Name = "MVb_Ay__Sy"
Option Explicit
Function SyAddAp(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Dim O$(), I
For Each I In Av
    If IsStr(I) Then
        Push O, I
    Else
        PushAy O, I
    End If
Next
End Function
Function Sy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Sy = AySy(Av)
End Function


Function SyEptStmt$(A)
Dim O$(), I
Push O, "Ept =  EmpSy"
For Each I In AyNz(A)
    Push O, FmtQQ("Push Ept, ""?""", Replace(I, """", """"""))
Next
SyEptStmt = JnCrLf(O)
End Function


Function SyShow(XX$, Sy$()) As String()
Dim O$()
Select Case Sz(Sy)
Case 0
    Push O, XX & "()"
Case 1
    Push O, XX & "(" & Sy(0) & ")"
Case Else
    Push O, XX & "("
    PushAy O, Sy
    Push O, XX & ")"
End Select
SyShow = O
End Function
