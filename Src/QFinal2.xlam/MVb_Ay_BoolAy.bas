Attribute VB_Name = "MVb_Ay_BoolAy"
Option Explicit

Function BoolAyAnd(A() As Boolean) As Boolean
BoolAyAnd = BoolAyIsAllTrue(A)
End Function

Function BoolAyIsAllFalse(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then Exit Function
Next
BoolAyIsAllFalse = True
End Function

Function BoolAyIsAllTrue(A() As Boolean) As Boolean
Dim I
For Each I In A
    If Not I Then Exit Function
Next
BoolAyIsAllTrue = True
End Function

Function BoolAyIsSomTrue(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then BoolAyIsSomTrue = True: Exit Function
Next
End Function

Function BoolAyOr(A() As Boolean) As Boolean
BoolAyOr = BoolAyIsSomTrue(A)
End Function


Function BoolAyIsSomFalse(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then BoolAyIsSomFalse = True: Exit Function
Next
End Function


Function BoolOp(BoolOpStr) As eBoolOp
Dim O As eBoolOp
Select Case UCase(BoolOpStr)
Case "AND": O = eBoolOp.eOpAND
Case "OR": O = eBoolOp.eOpOR
Case "EQ": O = eBoolOp.eOpEQ
Case "NE": O = eBoolOp.eOpNE
Case Else: Stop
End Select
BoolOp = O
End Function

Function BoolOpStr_IsAndOr(A$) As Boolean
Select Case UCase(A)
Case "AND", "OR": BoolOpStr_IsAndOr = True
End Select
End Function

Function BoolOpStr_IsEqNe(A$) As Boolean
Select Case UCase(A)
Case "EQ", "NE": BoolOpStr_IsEqNe = True
End Select
End Function

Function BoolOpStr_IsVdt(A$) As Boolean
BoolOpStr_IsVdt = IsInUCaseSy(A, BoolOpSy)
End Function

Function BoolTxt$(A As Boolean, T$)
If A Then BoolTxt = T
End Function

Function BoolOpSy() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SslSy("AND OR")
End If
BoolOpSy = Y
End Function
