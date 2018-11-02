Attribute VB_Name = "MVb_Str_Quote"
Option Explicit
Function Quote$(A, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & A & .S2
End With
End Function

Function QuoteAsVb$(A)
QuoteAsVb = """" & Replace(A, """", """""") & """"
End Function

Function QuoteDbl$(A$)
QuoteDbl = """" & A & """"
End Function

Function QuoteSng$(A)
QuoteSng = "'" & A & "'"
End Function

Function QuoteSqBkt$(A)
QuoteSqBkt = "[" & A & "]"
End Function

Function QuoteSqBktIfNeed$(A)
If IsSqBktNeed(A) Then
    QuoteSqBktIfNeed = "[" & A & "]"
Else
    QuoteSqBktIfNeed = A
End If
End Function
