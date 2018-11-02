Attribute VB_Name = "MVb_X_S1S2_Fmt"
Option Explicit
Private Function ZW%(LinesAy$(), Nm$)
ZW = Max(LinesAyWdt(LinesAy), Len(Nm))
End Function

Private Function ZW1%(A() As S1S2, Nm1$)
ZW1 = ZW(S1S2AySy1(A), Nm1)
End Function

Function S1S2AyFmt(A() As S1S2, Optional Nm1$, Optional Nm2$) As String()
If ZIsLines(A) Then
    Dim W1%: W1 = ZW1(A, Nm1)
    Dim W2%: W2 = ZW2(A, Nm2)
    Dim H$: H = WdtAyHdrLin(ApIntAy(W1, W2))
    S1S2AyFmt = ZFmt(A, H, W1, W2, Nm1, Nm2)
    Exit Function
End If
S1S2AyFmt = ZFmtX(A, Nm1, Nm2)
End Function
Private Function ZFmtX(A() As S1S2, Nm1$, Nm2$) As String()
If Sz(A) = 0 Then Exit Function
Dim S1$(), Sep$, S2$()
    S1 = AyAlignL(S1S2AySy1(A))
    S2 = S1S2AySy2(A)
Sep = ZFmtX1(S1)
PushIAy ZFmtX, ZFmtX2(Nm1, Nm2, Len(S1(0)), Sep)
Dim J%
For J = 0 To UB(A)
    PushI ZFmtX, S1(J) & Sep & S2(J)
Next
End Function
Private Function ZFmtX2(Nm1$, Nm2$, W%, Sep$) As String()
If Nm1 = "" And Nm2 = "" Then Exit Function
PushI ZFmtX2, AlignL(Nm1, W) & Sep & Nm2
End Function
Private Function ZFmtX1$(A$())
Dim J%
For J = 0 To UB(A)
    If HasSpc(A(J)) Then ZFmtX1 = " | ": Exit Function
Next
ZFmtX1 = " "
End Function
Private Function ZIsLines(A() As S1S2) As Boolean
Dim J&
ZIsLines = True
For J = 0 To UB(A)
    With A(J)
        If IsLines(.S1) Then Exit Function
        If IsLines(.S2) Then Exit Function
    End With
Next
ZIsLines = False
End Function
Private Function ZW2%(A() As S1S2, Nm2$)
ZW2 = ZW(S1S2AySy2(A), Nm2)
End Function

Private Function ZFmt(A() As S1S2, H$, W1%, W2%, Nm1$, Nm2$) As String()
PushIAy ZFmt, ZFmt1(H, Nm1, Nm2, W1, W2)
Dim I&
PushI ZFmt, H
For I = 0 To UB(A)
   PushIAy ZFmt, ZFmt2(A(I), W1, W2)
   PushI ZFmt, H
Next
End Function
Private Function ZFmt1(H$, Nm1$, Nm2$, W1%, W2%) As String()
If Nm1 = "" And Nm2 = "" Then Exit Function
PushI ZFmt1, H
PushI ZFmt1, "| " & AlignL(Nm1, W1) & " | " & AlignL(Nm2, W2) & " |"
End Function
Private Function ZFmt2(A As S1S2, W1%, W2%) As String()
Dim S1$(), S2$(), U%
S1 = SplitCrLf(A.S1)
S2 = SplitCrLf(A.S2)
U = Max(UB(S1), UB(S2))
S1 = ZFmt3(S1, U, W1)
S2 = ZFmt3(S2, U, W2)
Dim J&
For J = 0 To U
    PushI ZFmt2, "| " & S1(J) & " | " & S2(J) & " |"
Next
End Function

Private Function ZFmt3(Ay$(), U%, W%) As String()
ReDim Preserve Ay(U)
Dim I
For Each I In AyNz(Ay)
    PushI ZFmt3, AlignL(I, W)
Next
End Function


Function WdtAyHdrLin$(A%())
Dim O$(), W
For Each W In A
    Push O, StrDup("-", W + 2)
Next
WdtAyHdrLin = "|" + Join(O, "|") + "|"
End Function

Private Sub Z_S1S2AyFmt()
Dim A() As S1S2, Nm1$, Nm2$
'Nm1 = "AA": Nm2 = "BB": A = ZZ_S1S2Ay: GoSub Tst
'Nm1 = "":   Nm2 = "":   A = ZZ_S1S2Ay: GoSub Tst
Nm1 = "AA":   Nm2 = "BB":   A = ZZ_S1S2Ay1: GoSub Tst
Exit Sub
Tst:
    Act = S1S2AyFmt(A, Nm1, Nm2)
    Brw Act
    Return
End Sub
Private Sub X(O1$(), O2$())
Erase O1
Erase O2
Dim A1$, A2$
A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df":          GoSub X
A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub X
A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df":               GoSub X
A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df":           GoSub X
A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df":            GoSub X
Exit Sub
X:
    PushI O1, A1
    PushI O2, A2
    Return
End Sub
Private Function ZZ_S1S2Ay1() As S1S2()
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushObj ZZ_S1S2Ay1, S1S2(A1(J), A2(J))
Next
End Function

Private Function ZZ_S1S2Ay() As S1S2()
Dim A1$(), A2$(), J%
X A1, A2
For J = 0 To UB(A1)
    PushObj ZZ_S1S2Ay, S1S2(RplVBar(A1(J)), RplVBar(A2(J)))
Next
End Function
