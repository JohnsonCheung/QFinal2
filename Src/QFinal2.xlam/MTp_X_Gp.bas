Attribute VB_Name = "MTp_X_Gp"
Option Explicit
Function LyGp(Ly$()) As Gp
Set LyGp = Gp(LyLnxAy(Ly))
End Function
Function Gp(A() As Lnx) As Gp
Set Gp = New Gp
With Gp
    .LnxAy = A
End With
End Function


Function CvGp(A) As Gp
Set CvGp = A
End Function



Function GpLy(A As Gp) As String()
GpLy = LnxAyLy(A.LnxAy)
End Function

Function GpRmvRmk(A As Gp) As Gp
Dim B() As Lnx: B = A.LnxAy
Dim M As Lnx
Dim J&, O() As Lnx
For J = 0 To UB(B)
    M = B(J)
    If Not LinIsTpRmkLin(M.Lin) Then
        PushObj O, M
    End If
Next
Set GpRmvRmk = Gp(O)
End Function
