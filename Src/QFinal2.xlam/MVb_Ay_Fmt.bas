Attribute VB_Name = "MVb_Ay_Fmt"
Option Explicit
Function DrFmtss$(A, W%())
Dim U%, J%
U = UB(A)
If U = -1 Then Exit Function
ReDim O$(U)
For J = 0 To U - 1
    O(J) = AlignL(A(J), W%(J))
Next
O(U) = A(U)
DrFmtss = JnSpc(O)
End Function

Function DrFmtssCell(A) As String()
Dim O$(), J&, X
O = AyReSz(O, A)
For Each X In AyNz(A)
    O(J) = Fmtss(X)
    J = J + 1
Next
DrFmtssCell = O
End Function
