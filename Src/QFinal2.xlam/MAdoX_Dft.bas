Attribute VB_Name = "MAdoX_Dft"
Option Explicit
Function DftWsNy(WsNy0, Fx$) As String()
Dim O$()
    O = CvSy(WsNy0)
If Sz(O) = 0 Then
    DftWsNy = FxWsNy(Fx)
Else
    DftWsNy = O
End If
End Function
Function DftTny(Tny0, Fb$) As String()
Dim O$()
    O = CvSy(Tny0)
If Sz(O) = 0 Then
    DftTny = FbTny(Fb)
Else
    DftTny = O
End If
End Function
