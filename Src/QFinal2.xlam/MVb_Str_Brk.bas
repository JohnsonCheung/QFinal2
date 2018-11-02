Attribute VB_Name = "MVb_Str_Brk"
Option Explicit
Private Sub ZZ_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, A$
A = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(A, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub


Function Brk(A, Sep, Optional NoTrim As Boolean) As S1S2
Const CSub$ = "Brk"
Dim P&: P = InStr(A, Sep)
If P = 0 Then Er CSub, "{S} does not contains {Sep}", A, Sep
Set Brk = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Function Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then Set Brk1 = S1S2Trim(A, "", NoTrim): Exit Function
Set Brk1 = BrkAt(A, P, Sep, NoTrim)
End Function

Function Brk1Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Set Brk1Rev = S1S2Trim(A, "", NoTrim): Exit Function
Set Brk1Rev = BrkAt(A, P, Sep, NoTrim)
End Function

Private Sub Z_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, A$
A = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(A, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

Function Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk2 = Brk2__X(A, P, Sep, NoTrim)
End Function

Function Brk2__X(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk2__X = S1S2("", A)
    Else
        Set Brk2__X = S1S2("", Trim(A))
    End If
    Exit Function
End If
Set Brk2__X = BrkAt(A, P, Sep, NoTrim)
End Function

Sub Brk2Asg(A, Sep$, O1$, O2$)
Dim P%: P = InStr(A, Sep)
If P = 0 Then
    O1 = ""
    O2 = Trim(A)
Else
    O1 = Trim(Left(A, P - 1))
    O2 = Trim(Mid(A, P + 1))
End If
End Sub

Function Brk2Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
Set Brk2Rev = Brk2__X(A, P, Sep, NoTrim)
End Function

Sub BrkAsg(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
BrkAtAsg A, InStr(A, Sep), Sep, O1, O2, NoTrim
End Sub

Function BrkAt1(A, P&, Sep, NoTrim As Boolean) As S1S2

End Function

Function BrkAt(A, P&, Sep, NoTrim As Boolean) As S1S2
Dim S1$, S2$
S1 = Left(A, P - 1)
S2 = Mid(A, P + Len(Sep))
Set BrkAt = S1S2Trim(S1, S2, NoTrim)
End Function

Sub BrkAtAsg(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
If At = 0 Then
    MsgBrw "[Str] does not have [Sep].  @BrkAtAsg.", A, Sep
    Stop
    Exit Sub
End If
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Function BrkBoth(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then
    Set BrkBoth = S1S2Trim(A, A, NoTrim)
    Exit Function
End If
Set BrkBoth = BrkAt(A, P, Sep, NoTrim)
End Function

Function BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim S1$, S2$
Select Case L
Case 0:
Case 1
    S1 = QuoteStr
    S2 = QuoteStr
Case 2
    S1 = Left(QuoteStr, 1)
    S2 = Right(QuoteStr, 1)
Case Else
    If InStr(QuoteStr, "*") > 0 Then
        Set BrkQuote = Brk(QuoteStr, "*", NoTrim:=True)
        Exit Function
    End If
    Stop
End Select
Set BrkQuote = S1S2(S1, S2)
End Function

Sub BrkQuoteAsg(QuoteStr$, O1$, O2$)
S1S2Asg BrkQuote(QuoteStr), O1, O2
End Sub

Function BrkRev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & A & "] does not contains Sep[" & Sep & "]"
BrkRev = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Sub BrkS1Asg(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
BrkS1AtAsg A, InStr(A, Sep), Sep, O1, O2, NoTrim
End Sub

Sub BrkS1AtAsg(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
If At = 0 Then
    O1 = A
    O2 = ""
    Exit Sub
End If
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub
