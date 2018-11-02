Attribute VB_Name = "MVb_Str_Nm"
Option Explicit
Private Function NmNy__IsBrk(A, OPos%, OLen%) As Boolean
Stop '
End Function
Function NmNy(A) As String()
If Not IsNm(A) Then Stop
Dim P%, CPos%, CLen%
For P = 2 To Len(A)
    If NmNy__IsBrk(A, CPos, CLen) Then
        PushI NmNy, Mid(A, CPos, CLen)
    End If
Next
If CLen > 0 Then
    PushI NmNy, Mid(A, CPos, CLen)
End If
End Function
Function NmSeqNo%(A)
Dim B$: B = TakAftRev(A, "_")
If B = "" Then Exit Function
If Not IsNumeric(B) Then Exit Function
NmSeqNo = B
End Function
Function StrNy(A) As String()
Dim O$: O = RplPun(A)
Dim O1$(): O1 = AyWhSingleEle(SslSy(O))
Dim O2$()
Dim J%
For J = 0 To UB(O1)
    If Not IsDigit(FstChr(O1(J))) Then Push O2, O1(J)
Next
StrNy = O2
End Function

Sub DDNmBrkAsg(A, O1$, O2$, O3$)
Dim Ay$(): Ay = Split(A, ".")
Select Case Sz(Ay)
Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
Case Else: Stop
End Select
End Sub


Function DDNmThird$(A)
Dim Ay$(): Ay = Split(A, "."): If Sz(Ay) <> 3 Then Stop
DDNmThird = Ay(2)
End Function
