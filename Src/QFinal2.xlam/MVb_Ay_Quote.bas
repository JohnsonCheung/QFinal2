Attribute VB_Name = "MVb_Ay_Quote"
Option Explicit
Function AyQuote(A, QuoteStr$) As String()
If Sz(A) = 0 Then Exit Function
Dim U&: U = UB(A)
Dim O$()
    ReDim O(U)
    Dim J&
    Dim Q1$, Q2$
    With BrkQuote(QuoteStr)
        Q1 = .S1
        Q2 = .S2
    End With
    For J = 0 To U
        O(J) = Q1 & A(J) & Q2
    Next
AyQuote = O
End Function

Function AyQuoteDbl(A) As String()
AyQuoteDbl = AyQuote(A, """")
End Function

Function AyQuoteSng(A) As String()
AyQuoteSng = AyQuote(A, "'")
End Function

Function AyQuoteSqBkt(A) As String()
AyQuoteSqBkt = AyQuote(A, "[]")
End Function

Function AyQuoteSqBktCsv$(A)
AyQuoteSqBktCsv = JnComma(AyQuoteSqBkt(A))
End Function

Function AyQuoteSqBktIfNeed(A) As String()
Dim X
For Each X In AyNz(A)
    PushI AyQuoteSqBktIfNeed, QuoteSqBktIfNeed(CStr(X))
Next
End Function



Function AySqBktQuoteIfNeed(A) As String()
Dim X
For Each X In AyNz(A)
    PushI AySqBktQuoteIfNeed, QuoteSqBktIfNeed(X)
Next
End Function
