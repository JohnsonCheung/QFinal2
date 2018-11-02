Attribute VB_Name = "MVb_Er_Msg"
Option Explicit
Sub FunMsgAv_Brw(A, Msg$, Av())
AyBrw FunMsgAvLy(A, Msg, Av)
End Sub

Function FunMsgAvLy(A, Msg$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(Msg)
C = NyAvLy(CvSy(AyAdd(ApSy("Fun"), MsgNy(Msg))), CvAy(AyAdd(Array(A), Av)))
FunMsgAvLy = AyAdd(B, C)
End Function

Sub MsgAp_Brw(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_Brw Msg, Av
End Sub

Sub MsgAp_Dmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
AyDmp MsgAvLy(A, Av)
End Sub

Function MsgAp_Lin$(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAp_Lin = MsgAvLin(A, Av)
End Function

Function MsgAp_Ly(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgAp_Ly = MsgAvLy(A, Av)
End Function

Sub MsgApSclDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
Debug.Print MsgAv_Scl(A, Av)
End Sub

Function NmvLy(Nm$, V) As String()
Dim Ly$(): Ly = VarLy(V)
Dim J%, S$
PushI NmvLy, Nm & ": " & Ly(0)
S = Space(Len(Nm) + 2)
For J = 1 To UB(Ly)
    PushI NmvLy, S & Ly(J)
Next
End Function

Function NmvStr$(Nm$, V)
NmvStr = Nm & "=[" & VarStr(V) & "]"
End Function

Function NyAvLin$(A$(), Av())
Dim U&
U = UB(A)
If U = -1 Then Exit Function
Dim O$(), J%
For J = 0 To U
    Push O, NmvStr(A(J), Av(J))
Next
NyAvLin = Join(AyAddPfx(O, " | "))
End Function

Function NyAvLy(A$(), Av(), Optional Indent%) As String()
Dim W%, O$(), J%, A1$(), A2$()
W = AyWdt(A)
A1 = AyAlignL(A)
AyabSetSamMax A1, Av
For J = 0 To UB(A)
    PushAy O, NmvLy(A1(J), Av(J))
Next
NyAvLy = AyAddPfx(O, Space(Indent))
End Function

Function NyAvScl$(A$(), Av())
Dim O$(), J%, X, Y
X = A
Y = Av
AyabSetSamMax X, Y
For J = 0 To UB(X)
    Push O, RmvSqBkt(X(J)) & "=" & VarStr(Y(J))
Next
NyAvScl = JnSemiColon(O)
End Function

Function NyLin$(A$(), Av())
NyLin = NyAvLin(A, Av)
End Function

Function NyLy(Ny0, Av(), Optional Indent% = 4) As String()
NyLy = NyAvLy(CvNy(Ny0), Av, Indent)
End Function

Sub NyLyDmp(A, ParamArray Ap())
Dim Av(): Av = Ap
D NyLy(CvNy(A), Av, 0)
End Sub

Function NyScl$(A$(), Av())
NyScl = NyAvScl(A, Av)
End Function
Sub FunMsgAvLyDmp(A$, Msg$, Av())
D FunMsgAvLy(A, Msg, Av)
End Sub

Sub FunMsgAvLinDmp(A$, Msg$, Av())
D FunMsgAvLin(A, Msg, Av)
End Sub

Function FunMsgAvScl(A, Msg$, Av())
FunMsgAvScl = A & ";" & MsgAv_Scl(Msg, Av)
End Function

Sub FunMsgBrw(A$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
AyBrw FunMsgAvLy(A, Msg, Av)
End Sub

Sub FunMsg(A$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgAvLinDmp A, Msg, Av
End Sub

Function Msg(SqBktMacroStr$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
Msg = NyAvLy(MacroNy(SqBktMacroStr), Av)
End Function

Sub FunMsgDmp(A$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgAvLinDmp A, Msg, Av
End Sub

Function FunMsgAvLin$(Fun$, MacroStr$, Av())

End Function

Function FunMsgLin$(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgLin = FunMsgAvLin(Fun, Msg, Av)
End Function

Sub FunMsgLinDmp(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgAvLin(Fun, Msg, Av)
End Sub

Function FunMsgLy(A, Msg$, Av()) As String()
FunMsgLy = FunMsgAvLy(A, Msg, Av)
End Function

Sub FunMsgLyDmp(A, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgAvLy(A, Msg, Av)
End Sub

Sub FunMsgSclDmp(A, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
D FunMsgAvScl(A, Msg, Av)
End Sub

Function NmssAvLy(A$, Av()) As String()
NmssAvLy = NyAvLy(SslSy(A), Av)
End Function

Function NmssApLy(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
NmssApLy = NyAvLy(SslSy(A), Av)
End Function

Function MacroStrAvLy(A$, Av()) As String()
MacroStrAvLy = NyAvLy(MacroNy(A, Bkt:="[]"), Av)
End Function

Function FunMsgLines$(CSub$, MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
Stop
'ErMsgLines = ErMsgLinesByAv(CSub, MacroStr, Av)
End Function
Sub MsgAv_Brw(A$, Av())
AyBrw MsgAvLy(A, Av)
End Sub

Function MsgAvLin$(A$, Av())
Dim B$(), C$
C = NyLin(MsgNy(A), Av)
MsgAvLin = EnsSfxDot(A) & C
End Function

Function MsgAvLy(A$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(A)
C = AyAddTab(NyAvLy(MsgNy(A), Av))
MsgAvLy = AyAdd(B, C)
End Function

Function MsgAv_Scl$(A$, Av())
MsgAv_Scl = A & ";" & NyScl(MsgNy(A), Av)
End Function

Sub MsgBrw(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_Brw A, Av
End Sub

Sub MsgBrwStop(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_Brw A, Av
Stop
End Sub

Sub MsgDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
AyDmp MsgAvLy(A, Av)
End Sub

Function MsgLin$(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgLin = MsgAvLin(A, Av)
End Function

Function MsgLy(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgLy = MsgAvLy(A, Av)
End Function

Function MsgNy(A) As String()
Dim O$(), P%, J%
O = Split(A, "[")
AyShf O
For J = 0 To UB(O)
    P = InStr(O(J), "]")
    O(J) = "[" & Left(O(J), P)
Next
MsgNy = O
End Function

Function MsgScl$(A$, Av())
MsgScl = MsgAv_Scl(A, Av)
End Function

Sub MsgSclDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
D MsgScl(A, Av)
End Sub
Function VarStr$(V)
Select Case True
Case IsPrim(V):    VarStr = V
Case IsArray(V):   VarStr = AyLines(V)
Case IsNothing(V): VarStr = "*Nothing"
Case IsObject(V):  VarStr = "*Type[" & TypeName(V) & "]"
Case IsEmpty(V):   VarStr = "*Empty"
Case IsMissing(V): VarStr = "*Missing"
Case Else: Stop
End Select
End Function
Sub Warn(CSub$, SqBktMacroStr$, ParamArray Ap())

End Sub

Function VarLy(A) As String()
VarLy = SplitCrLf(VarLines(A))
End Function

Function VarLines$(A, Optional Lvl%)
Dim T$, S$, W%, I, O$(), Sep$
Select Case True
Case IsDic(A): VarLines = JnCrLf(DicFmt(CvDic(A)))
Case IsPrim(A): VarLines = A
Case IsLinesAy(A): VarLines = LinesAyLines(CvSy(A))
Case IsSy(A): VarLines = JnCrLf(A)
Case IsNothing(A): VarLines = "#Nothing"
Case IsEmpty(A): VarLines = "#Empty"
Case IsMissing(A): VarLines = "#Missing"
Case IsObject(A): VarLines = "#Obj(" & TypeName(A) & ")"
Case IsArray(A)
    If Sz(A) = 0 Then Exit Function
    For Each I In A
        PushI O, VarLines(I, Lvl + 1)
    Next
    If Lvl > 0 Then
        W = LinesAyWdt(O)
        Sep = LvlSep(Lvl)
        PushI O, StrDup(Sep, W)
    End If
    VarLines = JnCrLf(O)
Case Else
End Select
End Function

Function LvlSep$(Lvl%)
Select Case Lvl
Case 0: LvlSep = "."
Case 1: LvlSep = "-"
Case 2: LvlSep = "+"
Case 3: LvlSep = "="
Case 4: LvlSep = "*"
Case Else: LvlSep = Lvl
End Select
End Function
