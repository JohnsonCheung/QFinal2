Attribute VB_Name = "MVb_Align"
Option Explicit

Function AlignL$(A, W)
Dim L%: L = Len(A)
If L >= W Then
    AlignL = A
Else
    AlignL = A & Space(W - Len(A))
End If
End Function

Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function


Function StrAlignL$(S$, W, Optional ErIFmnotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "StrAlignL"
Dim L%: L = Len(S)
If L > W Then
    If ErIFmnotEnoughWdt Then
        Stop
        'Er CSub, "Len({S)) > {W}", S, W
    End If
    If DoNotCut Then
        StrAlignL = S
        Exit Function
    End If
End If

If W >= L Then
    StrAlignL = S & Space(W - L)
    Exit Function
End If
If W > 2 Then
    StrAlignL = Left(S, W - 2) + ".."
    Exit Function
End If
StrAlignL = Left(S, W)
End Function

Function LinesAyAlignLasLin(A$()) As String()
Dim W%: W = LinesAyWdt(A)
Dim Lines
For Each Lines In AyNz(A)
    PushI LinesAyAlignLasLin, LinesAlignLasLin(Lines, W)
Next
End Function
Function LinesAlignLasLin$(A, W%)
Stop '
End Function
Function LinesAlign$(A, W%)
Stop '
End Function
Function LinesAyAlign(A$()) As String()
Dim W%: W = LinesAyWdt(A)
Dim Lines
For Each Lines In AyNz(A)
    PushI LinesAyAlign, LinesAlign(Lines, W)
Next
End Function
Function VblAlign$(Vbl$, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%)
VblAlign = JnVBar(VblAlignAsLy(Vbl, Pfx, IdentOpt, Sfx, WdtOpt))
End Function

Function VblAlignAsLines$(Vbl$, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%)
VblAlignAsLines = JnCrLf(VblAlignAsLy(Vbl, Pfx, IdentOpt, Sfx, WdtOpt))
End Function
Function VblAlignAsLy(Vbl$, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%) As String()
Ass IsVbl(Vbl)
If IsEmp(Vbl) Then Exit Function
Dim Wdt%
    Dim W%
    W = VblWdt(Vbl)
    If W > WdtOpt Then
        Wdt = W
    Else
        Wdt = WdtOpt
    End If
Dim Ident%
    If Ident < 0 Then
        Ident = 0
    Else
        Ident = IdentOpt
    End If
Dim O$()
    Dim Ay$()
    Ay = SplitVBar(Vbl)
    Dim J%, A$, U&, S$, S1$, P$
    U = UB(Ay)
    P = IIf(Pfx <> "", Pfx & " ", "")
    S1 = Space(Ident)
    For J = 0 To U
        If J = 0 Then
'            S = AlignL(P, Ident, DoNotCut:=True)
        Else
            S = S1
        End If
'        A = S & AlignL(Ay(J), Wdt, ErIfNotEnoughWdt:=True)
        If J = U Then
            A = A & " " & Sfx
        End If
        Push O, A
    Next
VblAlignAsLy = O
End Function
