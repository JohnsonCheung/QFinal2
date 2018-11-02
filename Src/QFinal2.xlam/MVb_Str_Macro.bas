Attribute VB_Name = "MVb_Str_Macro"
Option Explicit
Function MacroNy(MacroStr$, Optional ExlBkt As Boolean, Optional Bkt$ = "[]") As String()
Dim Q1$, Q2$
With BrkQuote(Bkt)
    Q1 = .S1
    Q2 = .S2
End With
If Q1 = Q2 Then Stop
If Len(Q1) <> 1 Then Stop
If Len(Q2) <> 1 Then Stop
If Not HasSubStr(MacroStr, Q1) Then Exit Function

Dim Ay$(): Ay = Split(MacroStr, Q1)
Dim O$(), J%
For J = 1 To UB(Ay)
    Push O, TakBef(Ay(J), Q2)
Next
If Not ExlBkt Then
    O = AyAddPfxSfx(O, Q1, Q2)
End If
MacroNy = O
End Function
