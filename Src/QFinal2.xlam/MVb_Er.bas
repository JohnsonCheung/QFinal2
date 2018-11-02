Attribute VB_Name = "MVb_Er"
Option Explicit
'Calling this module functions will throw error
Sub ErWh(CSub$, Msg$, FF, ParamArray Ap())
Dim Av(): Av = Ap
Thow CSub, Msg, SslSy(FF), Av
End Sub
Sub Er(CSub$, SqBktMacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
Thow CSub, SqBktMacroStr, MacroNy(SqBktMacroStr), Av
End Sub
Private Sub Thow(CSub$, VblMsg$, Ny$(), Av())
AyBrw ItmAddAy( _
    "Subr-" & CSub & " : " & RplVBar(VblMsg), _
    NyAvLy(Ny, Av))
Stop
End Sub

Sub AssChk(Chk$())
If Sz(Chk) = 0 Then Exit Sub
AyBrw Chk
Stop
End Sub

Sub RaiseErr()
Err.Raise -1, , "Please check messages opened in notepad"
End Sub

