Attribute VB_Name = "MVb_Lin_Ssl"
Option Explicit
Function SslAy_Sy(A$()) As String()
Dim O$(), L
If Sz(A) = 0 Then Exit Function
For Each L In A
    PushAy O, SslSy(L)
Next
SslAy_Sy = O
End Function

Function SslHas(A, N) As Boolean
SslHas = AyHas(SslSy(A), N)
End Function

Function SslIx&(A, N)
SslIx = AyIx(SslSy(A), N)
End Function

Function SslJnComma$(Ssl)
SslJnComma = JnComma(SslSy(Ssl))
End Function

Function SslJnQuoteComma$(Ssl)
SslJnQuoteComma = JnComma(AyQuote(SslSy(Ssl), "'"))
End Function

Function SslSqBktCsv$(A)
Dim B$(), C$()
B = SslSy(A)
C = AyQuoteSqBkt(B)
SslSqBktCsv = JnComma(C)
End Function

Function SslSy(A) As String()
SslSy = SplitSpc(RplDblSpc(Trim(A)))
End Function
