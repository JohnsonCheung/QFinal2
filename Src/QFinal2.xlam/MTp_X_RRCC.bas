Attribute VB_Name = "MTp_X_RRCC"
Option Explicit

Function RRCCIsEmp(A As RRCC) As Boolean
RRCCIsEmp = True
With A
   If .R1 <= 0 Then Exit Function
   If .R2 <= 0 Then Exit Function
   If .R1 > .R2 Then Exit Function
End With
RRCCIsEmp = False
End Function

Function CvRRCC(A) As RRCC
Set CvRRCC = A
End Function
Function RRCC(R1, R2, C1, C2) As RRCC
Dim O As New RRCC
With O
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
Set RRCC = O
End Function
