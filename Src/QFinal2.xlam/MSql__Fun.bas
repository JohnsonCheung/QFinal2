Attribute VB_Name = "MSql__Fun"
Option Explicit
Public Sql_Shared As New Sql_Shared
Private Function QChr$(A)
Dim O$
Select Case True
Case IsStr(A): O = "'"
Case IsDate(A): O = "#"
Case IsEmpty(A), IsNull(A), IsNothing(A): Stop
QChr = O
End Select
End Function
Function VarAySqlQuote(Vy) As String()
Dim V
For Each V In Vy
    PushI VarAySqlQuote, VarSqlQuote(V)
Next
End Function

Function VarSqlQuote$(A)
Dim Q$
Q = QChr(A)
VarSqlQuote = Q & A & Q
End Function
Function SampleDb_DutyPrepare() As Database

End Function
