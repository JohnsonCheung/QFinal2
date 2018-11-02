Attribute VB_Name = "MVb_Str_Has"
Option Explicit
Function Has(A, SubStr) As Boolean
Has = InStr(A, SubStr) > 0
End Function

Function HasCrLf(A) As Boolean
HasCrLf = Has(A, vbCrLf)
End Function

Function HasDot(A) As Boolean
HasDot = InStr(A, ".") > 0
End Function

Function HasHyphen(A) As Boolean
HasHyphen = HasSubStr(A, "-")
End Function

Function HasPound(A) As Boolean
HasPound = InStr(A, "#") > 0
End Function

Function HasSpc(A) As Boolean
HasSpc = InStr(A, " ") > 0
End Function

Function HasSqBkt(A) As Boolean
HasSqBkt = FstChr(A) = "[" And LasChr(A) = "]"
End Function

Function HasSubStr(A, SubStr$, Optional Cmp As VbCompareMethod = vbTextCompare) As Boolean
HasSubStr = InStr(1, A, SubStr, Cmp) > 0
End Function

Function HasSubStrAy(A, SubStrAy$()) As Boolean
Dim S
For Each S In SubStrAy
    If HasSubStr(A, CStr(S)) Then HasSubStrAy = True: Exit Function
Next
End Function

Function HasT1(A$, T$) As Boolean
HasT1 = LinT1(A) = T
End Function

Function HasVBar(A$) As Boolean
HasVBar = HasSubStr(A, "|")
End Function
