Attribute VB_Name = "MDta_Fmt_Fmtss"
Option Explicit
Function DryFmtss(A()) As String()
Dim W%(), Dr, O$()
W = DryWdt(A)
For Each Dr In AyNz(A)
    PushI O, DrFmtss(Dr, W)
Next
DryFmtss = O
End Function

Function DrFmtss$(A, W%())
DrFmtss = DrFmt(A, W, " ")
End Function

Function AyColBrkssDry(A, ColBrkss$) As Variant()
Dim Lin, Ay$()
Ay = SslSy(ColBrkss)
For Each Lin In AyNz(A)
    PushI AyColBrkssDry, LinBrkssDr(Lin, Ay)
Next
End Function

Sub DrsFmtssDmp(A As Drs)
D DrsFmtss(A)
End Sub

Function DrsFmtss(A As Drs) As String()
DrsFmtss = DryFmtss(CvAy(ItmAddAy(A.Fny, A.Dry)))
End Function

Sub DrsFmtssBrw(A As Drs)
Brw DrsFmtss(A)
End Sub

Function DryFmtssCell(A()) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    Push DryFmtssCell, DrFmtssCell(Dr)
Next
End Function
Function LinBrkssDr(Lin, BrkssAy$()) As String()
Dim Brk, P%, L$
L = Lin
For Each Brk In BrkssAy
    P = InStr(L, Brk)
    If P = 0 Then Exit For
    Push LinBrkssDr, Left(L, P - 1)
    L = Mid(L, P)
Next
Push LinBrkssDr, L
End Function



Function AyFmt(A, ColBrkss$) As String()
AyFmt = DryFmtss(AyColBrkssDry(A, ColBrkss))
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
