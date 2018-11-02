Attribute VB_Name = "MIde__Identifier"
Option Explicit
Function MthLyExternalNy(A$()) As String()
Dim PmNy$(): PmNy = LinPmNy(SrcContLin(A, 0))
Dim DimNy$(): DimNy = MthLyDimNy(A)
MthLyExternalNy = AyMinusAp(StrIdentifierAy(JnSpc(A)), DimNy, PmNy)
End Function

Function MthLyDimNy(A$()) As String()
Dim S
For Each S In AyNz(MthLyDimStmtAy(A))
    PushIAy MthLyDimNy, DimStmtNy(S)
Next
End Function
Function DimStmtNy(A) As String()

End Function
Function MthLyDimStmtAy(A$()) As String()

End Function
Function StrIdentifierAy(A) As String()

End Function
Function StrNy(A) As String()
Dim O$, J%
O = A
Const C$ = "~!`@#$%^&*()-_=+[]{};;""'<>,.?/" & vbCr & vbLf
For J = 1 To Len(C)
    O = Replace(O, Mid(C, J, 1), " ")
Next
StrNy = AyWhDist(SslSy(O))
End Function
Function VbKwAy() As String()
Static X$()
If Sz(X) = 0 Then
    X = SslSy("Function Sub Then If As For To Each End While Wend Loop Do Static Dim Option Explicit Compare Text")
End If
VbKwAy = X
End Function
