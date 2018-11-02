Attribute VB_Name = "MVb_JnSplit_Split"
Option Explicit
Function SplitComma(A) As String()
SplitComma = Split(A, ",")
End Function

Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function

Function SplitDot(A) As String()
SplitDot = Split(A, ".")
End Function

Function SplitLf(A) As String()
SplitLf = Split(A, vbLf)
End Function

Function SplitLines(A) As String()
Dim B$: B = Replace(A, vbCrLf, vbLf)
SplitLines = SplitLf(B)
End Function

Function SplitSC(A) As String()
SplitSC = Split(A, ";")
End Function

Function SplitSpc(A) As String()
SplitSpc = Split(A, " ")
End Function

Function SplitSsl(A) As String()
SplitSsl = Split(RplDblSpc(Trim(A)), " ")
End Function

Function SplitVBar(A) As String()
SplitVBar = Split(A, "|")
End Function
