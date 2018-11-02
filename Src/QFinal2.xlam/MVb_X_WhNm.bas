Attribute VB_Name = "MVb_X_WhNm"
Option Explicit
Function WhNm(Optional Patn$, Optional Exl) As WhNm
Dim O As New WhNm
Set O.Re = Re(Patn)
O.ExlAy = CvSy(Exl)
Set WhNm = O
End Function
