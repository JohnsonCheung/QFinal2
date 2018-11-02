Attribute VB_Name = "MVb_JnSplit_Jn"
Option Explicit
Function Jn$(A, Sep$)
Jn = Join(AySy(A), Sep)
End Function

Function JnComma$(A)
JnComma = Jn(A, ",")
End Function

Function JnCommaCrLf$(A)
JnCommaCrLf = Jn(A, "," & vbCrLf)
End Function

Function JnAnd$(A)
JnAnd = Jn(A, " and ")
End Function

Function JnCommaSpc(A)
JnCommaSpc = Jn(A, ", ")
End Function

Function JnCrLf$(A)
JnCrLf = Jn(A, vbCrLf)
End Function

Function JnCrLfWithIx$(A)
JnCrLfWithIx = JnCrLf(AyAddIxPfx(A))
End Function

Function JnDblCrLf$(A)
JnDblCrLf = Jn(A, vbCrLf & vbCrLf)
End Function

Function JnDot$(A)
JnDot = Jn(A, ".")
End Function

Function JnQDblComma$(A)
JnQDblComma = JnComma(AyQuoteDbl(A))
End Function

Function JnQDblSpc$(A)
JnQDblSpc = JnSpc(AyQuoteDbl(A))
End Function

Function JnQSngComma$(A)
JnQSngComma = JnComma(AyQuoteSng(A))
End Function

Function JnQSngSpc$(A)
JnQSngSpc = JnSpc(AyQuoteSng(A))
End Function

Function JnQSqBktComma$(A)
JnQSqBktComma = JnComma(AyQuoteSqBkt(A))
End Function

Function JnQSqBktSpc$(A)
JnQSqBktSpc = JnSpc(AyQuoteSqBkt(AySy(A)))
End Function

Function JnSemiColon$(A)
JnSemiColon = Jn(A, ";")
End Function

Function JnSpc$(A)
JnSpc = Jn(A, " ")
End Function

Function JnTab$(A)
JnTab = Join(A, vbTab)
End Function

Function JnTerm$(A)
Dim O$(), X
For Each X In AyNz(A)
    PushI O, QuoteSqBktIfNeed(CStr(X))
Next
JnTerm = Join(O, " ")
End Function

Function JnVBar$(A)
JnVBar = Jn(A, "|")
End Function
