Attribute VB_Name = "MVb_X_FTNo"
Option Explicit

Function FTNoAyLinCnt%(A() As FTNo)
Dim O%, M
For Each M In A
    O = O + FTNoLinCnt(CvFTNo(M))
Next
End Function

Function FTNoLinCnt%(A As FTNo)
Dim O%
O = A.ToNo - A.FmNo + 1
If O < 0 Then Stop
FTNoLinCnt = O
End Function

Function FTNo(FmNo%, ToNo%) As FTNo
Dim O As New FTNo
Set FTNo = O.Init(FmNo, ToNo)
End Function
