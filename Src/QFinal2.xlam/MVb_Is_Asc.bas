Attribute VB_Name = "MVb_Is_Asc"
Option Explicit

Function AscIsDig(A%) As Boolean
AscIsDig = &H30 <= A And A <= &H39
End Function

Function AscIsDigit(A%) As Boolean
If A < 48 Then Exit Function
If A > 57 Then Exit Function
AscIsDigit = True
End Function

Function AscIsFstNmChr(A%) As Boolean
AscIsFstNmChr = AscIsLetter(A)
End Function

Function AscIsLCase(A%) As Boolean
If A < 97 Then Exit Function
If A > 122 Then Exit Function
AscIsLCase = True
End Function

Function AscIsLetter(A%) As Boolean
AscIsLetter = True
If AscIsUCase(A) Then Exit Function
If AscIsLCase(A) Then Exit Function
AscIsLetter = False
End Function

Function AscIsNmChr(A%) As Boolean
AscIsNmChr = True
If AscIsLetter(A) Then Exit Function
If AscIsDig(A) Then Exit Function
AscIsNmChr = A = 95 '_
End Function

Function AscIsUCase(A%) As Boolean
If A < 65 Then Exit Function
If A > 90 Then Exit Function
AscIsUCase = True
End Function
