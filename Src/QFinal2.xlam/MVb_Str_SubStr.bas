Attribute VB_Name = "MVb_Str_SubStr"
Option Explicit

Function LasChr$(A)
LasChr = Right(A, 1)
End Function

Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function FstTwoChr$(A)
FstTwoChr = Left(A, 2)
End Function
