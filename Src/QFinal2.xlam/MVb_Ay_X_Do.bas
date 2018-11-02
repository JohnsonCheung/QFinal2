Attribute VB_Name = "MVb_Ay_X_Do"
Option Explicit
Sub AyDo(A, FunNm$)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Run FunNm, I
Next
End Sub

Sub AyDoABX(Ay, ABX$, A, B)
Dim X
For Each X In AyNz(Ay)
    Run ABX, A, B, X
Next
End Sub

Sub AyDoAXB(Ay, AXB$, A, B)
Dim X
For Each X In AyNz(Ay)
    Run AXB, A, X, B
Next
End Sub

Sub AyDoPPXP(A, PPXP$, P1, P2, P3)
Dim X
For Each X In AyNz(A)
    Run PPXP, P1, P2, X, P3
Next
End Sub

Sub AyDoPX(A, PX$, P)
Dim X
For Each X In AyNz(A)
    Run PX, P, X
Next
End Sub

Sub AyDoXP(A, XP$, P)
Dim X
For Each X In AyNz(A)
    Run XP, X, P
Next
End Sub
