Attribute VB_Name = "MVb_Lin_Term"
Option Explicit
Const CMod$ = "MVb_Lin_Term."

Function LinTermAy(A) As String()
Dim L$, J%
L = A
While L <> ""
    J = J + 1: If J > 50000 Then Stop
    PushI LinTermAy, ShfTerm(L)
Wend
End Function

Function ShfT$(O)
ShfT = ShfTerm(O)
End Function

Private Function ShfTerm1$(O)
Const CSub$ = CMod & "ShfTerm1"
Dim A$
A = LTrim(O)
Dim P%
P = InStr(A, "]")
If P = 0 Then Er CSub, "Given [Str] has Opn-Sq-Bkt but no Cls-Sq-Bkt", O
ShfTerm1 = Mid(A, 2, P - 2)
ShfTerm1 = LTrim(Mid(A, P + 1))
End Function

Function ShfTerm$(O)
Dim A$
    A = LTrim(O)
If FstChr(A) = "[" Then ShfTerm = ShfTerm1(O): Exit Function
Dim P%
    P = InStr(A, " ")
If P = 0 Then
    ShfTerm = A
    O = ""
    Exit Function
End If
ShfTerm = Left(A, P - 1)
O = LTrim(Mid(A, P + 1))
End Function

Private Sub Z_ShfT()
Dim O$, OEpt$
O = " S   DFKDF SLDF  "
OEpt = "DFKDF SLDF  "
Ept = "S"
GoSub Tst
'
O = " AA BB "
Ept = "AA"
OEpt = "BB "
GoSub Tst
'
Exit Sub
Tst:
    Act = ShfT(O)
    C
    Ass O = OEpt
    Return
End Sub
