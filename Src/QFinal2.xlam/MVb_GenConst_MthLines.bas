Attribute VB_Name = "MVb_GenConst_MthLines"
Option Explicit

Function ConstValMthLInes$(ConstVal$, Nm$, Optional IsPub As Boolean) _
' Return [MthLines] by [ConstVal$] and [Nm$]
If ConstVal = "" Then Stop
Dim A$()
Dim NChunk%
    A = SplitCrLf(ConstVal)
    NChunk = ZNChunk(Sz(A))
Dim O$()
    Dim J%
    For J = 0 To NChunk - 1
        PushI O, ZChunk(A, J)
    Next
    PushI O, ZLasLin(Nm, NChunk)
ConstValMthLInes = ZMakeMth(JnCrLf(O), Nm, IsPub)
End Function

Private Function ZMakeMth$(Lines$, Nm$, IsPub As Boolean)
Dim L1$, L2$
L1 = IIf(IsPub, "", "Private ") & "Function " & Nm & "$()" & vbCrLf
L2 = vbCrLf & "End Function"
ZMakeMth = vbCrLf & L1 & Lines & L2
End Function

Private Function ZNChunk%(Sz%)
ZNChunk = ((Sz - 1) \ 20) + 1
End Function

Function ZLasLin$(Nm$, NChunk%)
Dim B$
    Dim O$(), J%
    For J = 1 To NChunk
        PushI O, "A_" & J
    Next
    B = Join(O, " & vbCrLf & ")
ZLasLin = Nm & " = " & B
End Function

Private Function ZChunk$(ConstLy$(), IChunk%)
If Sz(ConstLy) = 0 Then Stop
Dim Ly$()
    Ly = AyMid(ConstLy, IChunk * 20, 20)
Dim O$()
    Dim L$, J&, U&
    U = UB(Ly)
    For J = 0 To U
        L = QuoteAsVb(Ly(J))
        Select Case True
        Case J = 0 And J = U: Push O, FmtQQ("Const A_?$ = ?", IChunk + 1, L)
        Case J = 0:           Push O, FmtQQ("Const A_?$ = ? & _", IChunk + 1, L)
        Case J = U:           Push O, "vbCrLf & " & L
        Case Else:            Push O, "vbCrLf & " & L & " & _"
        End Select
    Next
ZChunk = JnCrLf(O)
End Function

Private Sub Z_ConstValMthLines()
Dim Nm$, ConstVal$
GoSub Tst1
'GoSub Tst2
'--
Tst1:
    Nm = "ZZ_A"
    With Application.Vbe.VBProjects("QVb").VBComponents("M_GenConst_MthLines").CodeModule
        ConstVal = .Lines(1, .CountOfLines)
    End With
    GoSub Tst
    Return
'
Tst2:
    Nm = "ZZ_A"
    ConstVal = "AAA"
    GoSub Tst
    Return
Exit Sub
Tst:
    Act = ConstValMthLInes(ConstVal, Nm)
    Brw Act
    Stop
    C
    Return
End Sub

Sub Z()
Z_ConstValMthLines
End Sub

