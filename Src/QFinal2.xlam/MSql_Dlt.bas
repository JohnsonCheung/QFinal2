Attribute VB_Name = "MSql_Dlt"
Option Explicit
Function QDlt$(T, Optional BExpr$)
'QDlt = "Delete * from [" & T & "]" & X.WhBExprSqp(BExpr)
End Function

Function DltInAySqy(T, F, InAy, QChr$) As String()
Const C$ = "Delete * from [?]?"
Dim X, O$(), Q$, L%, M%
For Each X In InAy
    PushI O, Q & X & Q
    L = L + Len(X) + 3
    If L > 3000 Then
    Else
    End If
Next
If Sz(O) > 0 Then
    
End If
End Function
