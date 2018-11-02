VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sql_Shared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const CMod$ = "Sql_Shared."
Const C_Upd$ = "update"
Const C_Into$ = "into"
Const C_Sel$ = "select"
Const C_SelDis$ = "select distinct"
Const C_Fm$ = "from"
Const C_Gp$ = "group by"
Const C_Wh$ = "where"
Const C_And$ = "and"
Const C_Jn$ = "join"
Const C_LeftJn$ = "left join"
Const C_NLT$ = vbCrLf & vbTab
Const C_NLTT$ = vbCrLf & vbTab & vbTab
Public IsFmt As Boolean
Sub Z()
Z_SetSqpFmt
Z_WhFldInAy
Z_WhFldInAySqpAy
Z_WhFnyEqAy
End Sub
Private Function NL$() ' New Line
If IsFmt Then
    NL = vbCrLf
Else
    NL = " "
End If
End Function
Private Function NLT$() ' New Line Tabe
If IsFmt Then
    NLT = C_NLT
Else
    NLT = " "
End If
End Function
Private Function NLTT$() ' New Line Tabe
If IsFmt Then
    NLTT = C_NLTT
Else
    NLTT = " "
End If
End Function

Private Function AyQ(Ay) As String()
AyQ = AyQuote(Ay, VarSqlQuote(Ay(0)))
End Function
Function FldInAy$(F, InAy)
FldInAy = Q(F) & "(" & JnComma(AyQ(InAy)) & ")"
End Function
Private Function EyAsFnyLines$(Ey$(), Fny$())
Stop '
End Function
Function FFJnComma$(FF)
FFJnComma = JnComma(FFFny(FF))
End Function
Friend Function XInto$(T)
XInto = NLT & "Into" & NLTT & "[" & T & "]"
End Function

Function AddColSqp$(Fny0, FldDfnDic As Dictionary)
Dim Fny$(), O$(), J%
Fny = CvNy(Fny0)
ReDim O(UB(Fny))
For J = 0 To UB(Fny)
    O(J) = Fny(J) & " " & FldDfnDic(Fny(J))
Next
'AddColSqp = NxtLin & "Add Column " & JnComma(O)
End Function

Function AndSqp$(Expr$)
If Expr = "" Then Exit Function
'AndSqp = NxtLin & "and " & NxtLinTab & Expr
End Function

Private Function AyNxtLinTabTab$(A)
'AyNxtLinTabTab = vbCrLf & JnCrLf(AyAddPfx(A, Pfx & Pfx))
End Function

Function SqpExprIn$(Expr$, InLis$)
If InLis = "" Then Exit Function
SqpExprIn = FmtQQ("? in (?)", Expr, InLis)
End Function

Function Fm$(T)
Fm = NLT & C_Fm & NLTT & QuoteSqBkt(T)
End Function

Function Gp$(ExprVblAy$())
Ass IsVblAy(ExprVblAy)
Gp = VblAy_AlignAsLines(ExprVblAy, "|  Group By")
End Function

Sub Z_Gp()
Dim ExprVblAy$()
    Push ExprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push ExprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push ExprVblAy, "3sldkfjsdf"
AyDmp SplitVBar(Gp(ExprVblAy))
End Sub

Function SqpSelFldLvs$(FldLvs$, ExprVblAy$())
Dim Fny$(): Fny = SslSy(FldLvs)
'SqpSelFldLvs = SqpSel(Fny, ExprVblAy)
End Function

Private Sub SqpSel()
Dim Fny$(), ExprVblAy$()
ExprVblAy = ApSy("F1-Expr", "F2-Expr   AA|BB    X|DD       Y", "F3-Expr  x")
Fny = SplitSpc("F1 F2 F3xxxxx")
'Debug.Print RplVBar(SqpSel(Fny, ExprVblAy))
End Sub

Function Into$(T)
Into = NLT & C_Into & NLTT & Q(T)
End Function
Function SelX$(X, Optional IsDis As Boolean)
SelX = C_Sel & Dis(IsDis) & NLT & X
End Function
Private Function Dis$(IsDis As Boolean)
If IsDis Then Dis = " distinct"
End Function
Function Sel$(FF, Optional IsDis As Boolean)
Sel = SelX(FFJnComma(FF), IsDis)
End Function

Function SetFFEy$(FF, Ey$())
Const CSub$ = CMod & "SetFFEy"
Dim Fny$(): Fny = SslSy(FF)
Ass IsVblAy(Ey)
If Sz(Fny) <> Sz(Ey) Then Er CSub, "[FF-Sz} <> [Sz-Ey], where [FF],[Ey]", Sz(Fny), Sz(Ey), FF, Ey
Dim AFny$()
    AFny = AyAlignL(Fny)
    AFny = AyAddSfx(AFny, " = ")
Dim W%
    W = VblAyWdt(Ey)
Dim Ident%
    W = AyWdt(AFny)
Dim Ay$()
    Dim J%, U%, S$
    U = UB(AFny)
    For J = 0 To U
        If J = U Then
            S = ""
        Else
            S = ","
        End If
        Push Ay, VblAlign(Ey(J), Pfx:=AFny(J), IdentOpt:=Ident, WdtOpt:=W, Sfx:=S)
    Next
Dim Vbl$
    Dim Ay1$()
    Dim P$
    For J = 0 To U
        If J = 0 Then P = "|  Set" Else P = ""
        Push Ay1, VblAlign(Ay(J), Pfx:=P, IdentOpt:=6)
    Next
    Vbl = JnVBar(Ay1)
SetFFEy = Vbl
End Function

Private Sub Z_SetFld()
Dim Fny$(), ExprVblAy$()
Fny = SslSy("a b c d")
Push ExprVblAy, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "2sdfkl|lskdfjdf| sdf"
Push ExprVblAy, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "4sf| sdf"
    Act = SetFld(Fny, ExprVblAy)
Debug.Print RplVBar(Act)
End Sub
Function SetFld$(Fny$(), Ay)

End Function
Private Function Q$(A)
Q = QuoteSqBktIfNeed(A)
End Function
Function Upd$(T)
Upd = C_Upd & NLT & Q(T)
End Function

Function Whfv(F, V) ' Ssk is single-Sk-value
Whfv = CWh & Q(F) & "=" & QV(V)
End Function

Function WhK$(K&, T)
WhK = Whfv(T & "Id", K)
End Function

Function WhBet$(F, FmV, ToV)
WhBet = CWh & Q(F) & QV(FmV) & CAnd & QV(ToV)
End Function

Private Function QV$(V)
QV = VarSqlQuote(V)
End Function
Private Function CAnd$()
CAnd = " " & C_And & " "
End Function
Private Function CWh$()
CWh = NLT & C_Wh & NLTT
End Function
Function Wh$(BExpr$)
If BExpr = "" Then Exit Function
Wh = CWh & BExpr
End Function
Function WhFldInAy$(F, InAy)
WhFldInAy = CWh & FldInAy(F, InAy)
End Function
Private Sub Z_WhFldInAy()
Dim Fny$(), Ay()
Fny = SslSy("A B C")
Ay = Array(1, "2", #2/1/2017#)
Ept = " where A=1 and B='2' and C=#2017-2-1#"
GoSub Tst
Exit Sub
Tst:
    Act = WhFldInAy(Fny, Ay)
    C
    Return
End Sub
Private Function FnyEqAy$(Fny$(), Ay)

End Function

Function WhFnyEqAy$(Fny$(), Ay)
WhFnyEqAy$ = CWh & FnyEqAy(Fny, Ay)
End Function


Function WhSqp$(WhBExpr$)
If WhBExpr = "" Then Exit Function
'WhSqp = X.NL & "Where" & NxtLinTab & WhBExpr
End Function

Private Sub Z_SetSqpFmt()
Dim Fny$(), Vy()
Ept = RplVBar("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Fny = LinTermAy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
    Act = SetSqp(Fny, Vy)
    C
    Return
End Sub

Private Sub Z_WhFldInAySqpAy()

End Sub

Function VblAy_AlignAsLines$(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAy, Optional Sep$ = ",")
VblAy_AlignAsLines = JnVBar(VblAy_AlignAsLy(ExprVblAy, Pfx, IdentOpt, SfxAy, Sep))
End Function

Function VblAy_AlignAsLy(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAyOpt, Optional Sep$ = ",") As String()
Dim NoSfxAy As Boolean
Dim SfxWdt%
Dim SfxAy$()
    NoSfxAy = IsEmp(SfxAy)
    If Not NoSfxAy Then
        Ass IsSy(SfxAyOpt)
        SfxAy = AyAlignL(SfxAyOpt)
        Dim U%, J%: U = UB(SfxAy)
        For J = 0 To U
            If J <> U Then
                SfxAy(J) = SfxAy(J) & Sep
            End If
        Next
    End If
Ass IsVblAy(ExprVblAy)
Dim Ident%
    If IdentOpt > 0 Then
        Ident = IdentOpt
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$
U = UB(ExprVblAy)
Dim W%
    W = VblAyWdt(ExprVblAy)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If NoSfxAy Then
        If J = U Then S = "" Else S = Sep
    Else
        If J = U Then S = SfxAy(J) Else S = SfxAy(J) & Sep
    End If
    Push O, VblAlign(ExprVblAy(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
VblAy_AlignAsLy = O
End Function

Function QSel_Fny_EDic$(Fny$(), EDic As Dictionary, Optional IsDis As Boolean)
QSel_Fny_EDic = QSel_Fny_Ey(Fny, DicKySy(EDic, Fny), IsDis)
End Function

Function QSel_Fny_Ey$(Fny$(), Ey$(), Optional IsDis As Boolean)
QSel_Fny_Ey = SelX(EyAsFnyLines(Ey, Fny), IsDis)
End Function

Private Function FyEyAsLines(Fy$(), Ey$())
Dim O$()
    If IsFmt Then
        Dim Fy1$(), Ey1$(), J%
        Fy1 = AyAlignL(Fy)
        Ey1 = LinesAyAlignLasLin(Ey)
        For J = 0 To UB(Fy)
            PushI O, Ey1(J) & " as " & Fy1(J)
        Next
        FyEyAsLines = JnCommaCrLf(AyAddPfx(O, vbTab & vbTab))
        Return
    End If

For J = 0 To UB(Fy)
    If Fy(J) = Ey(J) Then
        PushI O, Fy(J)
    Else
        PushI O, Replace(Ey(J), vbCrLf, " ") & " as " & Fy(J)
    End If
Next
FyEyAsLines = JnCommaSpc(O)
End Function
Private Function FnyFy(A$()) As String()
FnyFy = AyQuoteSqBktIfNeed(A)
End Function
Function FnyEyAsLines$(Fny$(), Ey$())
FnyEyAsLines = FyEyAsLines(FnyFy(Fny), Ey)
End Function

Private Function FFFny(FF) As String()
FFFny = LinTermAy(FF)
End Function

Private Function FFEDicAsLines$(FF, E As Dictionary)
Dim Fny$(): Fny = FFFny(FF)
Dim Ey$(): Ey = DicKySy(E, Fny)
FFEDicAsLines = FnyEyAsLines(Fny, Ey)
End Function
Private Sub Z_WhFnyEqAy()

End Sub



