Attribute VB_Name = "MVb_Str_Fmt"
Option Explicit
Function FmtMacro$(MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtMacro = FmtMacroAv(MacroStr, Av)
End Function

Function FmtMacroAv$(MacroStr$, Av())
Dim Ay$(): Ay = MacroNy(MacroStr)
Dim O$: O = MacroStr
Dim J%, I
For Each I In Ay
    O = Replace(O, I, Av(J))
    J = J + 1
Next
FmtMacroAv = O
End Function

Function FmtMacroDic$(MacroStr$, Dic As Dictionary)
Dim O$: O = MacroStr
Dim I, K$
For Each I In AyNz(MacroNy(MacroStr))
    K = RmvFstLasChr(I)
    If Dic.Exists(K) Then
        O = Replace(O, I, Dic(K))
    End If
Next
FmtMacroDic = O
End Function

Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function FmtQQAv$(QQVbl, Av)
Dim O$, I, Cnt
O = Replace(QQVbl, "|", vbCrLf)
Cnt = SubStrCnt(QQVbl, "?")
If Cnt <> Sz(Av) Then
    MsgBrw "[QQVal] has [N-?], but not match with [Av]-[Sz]", QQVbl, Cnt, Av, Sz(Av)
    Stop
    Exit Function
End If
Dim P&
P = 1
For Each I In Av
    P = InStr(P, O, "?")
    If P = 0 Then Stop
    O = Left(O, P - 1) & Replace(O, "?", I, Start:=P, Count:=1)
    P = P + Len(I)
Next
FmtQQAv = O
End Function


Function Fmtss$(A)
Fmtss = QuoteSqBktIfNeed(EscSqBkt(EscCrLf(EscBackSlash(A))))
End Function

Function UnFmtss$(A)
UnFmtss = UnEscBackSlash(UnEscSqBkt(UnEscCrLf(A)))
End Function

Private Sub ZZ_FmtQQAv()
Debug.Print FmtQQ("klsdf?sdf?dsklf", 2, 1)
End Sub


Function LblTabAyFmt(Lbl$, Ay) As String()
PushI LblTabAyFmt, Lbl
PushIAy LblTabAyFmt, AyAddTab(Ay)
End Function
