Attribute VB_Name = "MVb__Bkt"
Option Explicit
Const CMod$ = "MVb_Str_Brk."
Sub Z()
Z_BrkBkt
Z_ZBktPosAsg
End Sub

Private Sub Z_ZBktPosAsg()
Dim A$, EptFmPos%, EptToPos%
'
A = "(A(B)A)A"
EptFmPos = 1
EptToPos = 7
GoSub Tst
'
A = " (A(B)A)A"
EptFmPos = 2
EptToPos = 8
GoSub Tst
'
A = " (A(B)A )A"
EptFmPos = 2
EptToPos = 9
GoSub Tst
'
Exit Sub
Tst:
    Dim ActFmPos%, ActToPos%
    ZBktPosAsg A, "(", ActFmPos, ActToPos
    Ass IsEq(ActFmPos, EptFmPos)
    Ass IsEq(ActToPos, EptToPos)
    Return
End Sub

Private Sub Z_BrkBkt()
Dim A$, OpnBkt$
A = "aaaa((a),(b))xxx":    OpnBkt = "(":          Ept = ApSy("aaaa", "(a),(b)", "xxx"): GoSub Tst
Exit Sub
Tst:
    Act = BrkBkt(A, OpnBkt)
    C
    Return
End Sub

Sub ZBktPosAsg(A, OpnBkt$, OFmPos%, OToPos%)
Const CSub$ = CMod & "ZBktPosAsg"
OFmPos = 0
OToPos = 0
'-- OFmPos
    Dim Q1$, Q2$
    Q1 = OpnBkt
    Q2 = ZClsBkt(OpnBkt)

    OFmPos = InStr(A, Q1)
    If OFmPos = 0 Then Exit Sub
'-- OToPos
    Dim NOpn%, J%
    For J = OFmPos + 1 To Len(A)
        Select Case Mid(A, J, 1)
        Case Q2
            If NOpn = 0 Then
                OToPos = J
                Exit For
            End If
            NOpn = NOpn - 1
        Case Q1
            NOpn = NOpn + 1
        End Select
    Next
    If OToPos = 0 Then Er CSub, "The bracket-[Q1]-[Q2] in [Str] has is not in pair: [Q1-Pos] is found, but not Q2-pos is 0", Q1, Q2, A, OFmPos
End Sub

Private Function ZClsBkt$(OpnBkt$)
Select Case OpnBkt
Case "(": ZClsBkt = ")"
Case "[": ZClsBkt = "]"
Case "{": ZClsBkt = "}"
Case Else: Stop
End Select
End Function

Private Function BrkBkt(A, Optional OpnBkt$ = vbOpnBkt) As String()
Dim P1%, P2%
    ZBktPosAsg A, OpnBkt, _
    P1, P2
Dim A1$, A2$, A3$
A1 = Left(A, P1 - 1)
A2 = Mid(A, P1 + 1, P2 - P1 - 1)
A3 = Mid(A, P2 + 1)
BrkBkt = ApSy(A1, A2, A3)
End Function

Function TakBetBkt$(A, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
ZBktPosAsg A, OpnBkt, P1, P2
TakBetBkt = Mid(A, P1 + 1, P2 - P1 - 1)
End Function

Function TakAftBkt$(A, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
   ZBktPosAsg A, OpnBkt, P1, P2
If P2 = 0 Then Exit Function
TakAftBkt = Mid(A, P2 + 1)
End Function

Function TakBefBkt$(A, Optional OpnBkt$ = vbOpnBkt)
Dim P1%, P2%
   ZBktPosAsg A, OpnBkt, P1, P2
If P1 = 0 Then Exit Function
TakBefBkt = Left(A, P1 - 1)
End Function

