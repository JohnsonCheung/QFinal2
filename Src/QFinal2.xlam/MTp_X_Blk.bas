Attribute VB_Name = "MTp_X_Blk"
Option Explicit
Function Blk(BlkTyStr$, A As Gp) As Blk
Set Blk = New Blk
With Blk
    .BlkTyStr = BlkTyStr
    Set .Gp = A
End With
End Function


Function CvBlk(A) As Blk
Set CvBlk = A
End Function
