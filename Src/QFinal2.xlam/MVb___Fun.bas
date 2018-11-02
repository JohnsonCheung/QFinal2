Attribute VB_Name = "MVb___Fun"
Option Explicit

Sub Asg(Fm, OTo)
If IsObject(Fm) Then
    Set OTo = Fm
Else
    If IsNull(Fm) Then
        OTo = ""
    Else
        OTo = Fm
    End If
End If
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub

Sub FunTim(Fun$)
Dim A!, B!
A = Timer
Run Fun
B = Timer
Debug.Print Fun, B - A
End Sub

Function InstrN&(S, SubStr, N%)
Dim P&, J%
For J = 1 To N
    P = InStr(P + 1, S, SubStr)
    If P = 0 Then Exit Function
Next
InstrN = P
End Function

Private Sub Z_InstrN()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InstrN(S, SubStr, N)
Ass Exp = Act
End Sub

Function Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) > O Then O = Av(J)
Next
Max = O
End Function
Function MaxVbTy(A As VbVarType, B As VbVarType) As VbVarType
If A = vbString Or B = vbString Then MaxVbTy = A: Exit Function
If A = vbEmpty Then MaxVbTy = B: Exit Function
If B = vbEmpty Then MaxVbTy = A: Exit Function
If A = B Then MaxVbTy = A: Exit Function
Dim AIsNum As Boolean, BIsNum As Boolean
AIsNum = IsVbTyNum(A)
BIsNum = IsVbTyNum(B)
Select Case True
Case A = vbBoolean And BIsNum: MaxVbTy = B
Case AIsNum And B = vbBoolean: MaxVbTy = A
Case A = vbDate Or B = vbDate: MaxVbTy = vbString
Case AIsNum And BIsNum:
    Select Case True
    Case A = vbByte: MaxVbTy = B
    Case B = vbByte: MaxVbTy = A
    Case A = vbInteger: MaxVbTy = B
    Case B = vbInteger: MaxVbTy = A
    Case A = vbLong: MaxVbTy = B
    Case B = vbLong: MaxVbTy = A
    Case A = vbSingle: MaxVbTy = B
    Case B = vbSingle: MaxVbTy = A
    Case A = vbDouble: MaxVbTy = B
    Case B = vbDouble: MaxVbTy = A
    Case A = vbCurrency Or B = vbCurrency: MaxVbTy = A
    Case Else: Stop
    End Select
Case Else: Stop
End Select
End Function

Function CanCvLng(A) As Boolean
On Error GoTo X
Dim L&: L = CLng(A)
CanCvLng = True
X:
End Function

Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
Min = AyMin(Av)
End Function

Sub SndKeys(A$)
DoEvents
SendKeys A, True
End Sub

Sub Brw(A)
Select Case True
Case IsStr(A): StrBrw A
Case IsArray(A): AyBrw A
Case Else: Stop
End Select
End Sub

Function Re(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
If Patn = "" Or Patn = "." Then Exit Function
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set Re = O
End Function
