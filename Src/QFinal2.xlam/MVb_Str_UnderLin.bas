Attribute VB_Name = "MVb_Str_UnderLin"
Option Explicit
Function UnderLin$(A$)
UnderLin = String(Len(A), "-")
End Function

Function UnderLinDbl$(A)
UnderLinDbl = String(Len(A), "=")
End Function

Function PushMsgUnderLinDbl(O$(), M$)
Push O, M
Push O, UnderLinDbl(M)
End Function

Function PushUnderLin(O$())
Push O, UnderLin(AyLasEle(O))
End Function

Function PushUnderLinDbl(O$())
Push O, UnderLinDbl(AyLasEle(O))
End Function

Function LinesUnderLin$(A)
LinesUnderLin = StrDup("-", LinesWdt(A))
End Function

Function PushMsgUnderLin(O$(), M$)
Push O, M
Push O, UnderLin(M)
End Function
