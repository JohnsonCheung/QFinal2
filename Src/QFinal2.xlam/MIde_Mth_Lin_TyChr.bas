Attribute VB_Name = "MIde_Mth_Lin_TyChr"
Option Explicit
Const TyChrLis$ = "!@#$%^&"

Function IsTyChr(A$) As Boolean
If Len(A) <> 1 Then Exit Function
IsTyChr = HasSubStr(TyChrLis, A)
End Function


Function TyChrAsTyStr$(TyChr$)
Dim O$
Select Case TyChr
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Stop
End Select
TyChrAsTyStr = O
End Function
