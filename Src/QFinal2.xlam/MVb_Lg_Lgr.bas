Attribute VB_Name = "MVb_Lg_Lgr"
Option Explicit

Sub LgrBrw()
FtBrw LgrFt
End Sub

Function LgrFilNo%()
LgrFilNo = FtAppFilNo(LgrFt)
End Function

Function LgrFt$()
LgrFt = LgrPth & "Log.txt"
End Function

Sub LgrLg(Msg$)
Dim F%: F = LgrFilNo
Print #F, NowStr & " " & Msg
If LgrFilNo = 0 Then Close #F
End Sub

Function LgrPth$()
Dim O$:
'O = WrkPth: PthEns O
O = O & "Log\": PthEns O
LgrPth = O
End Function
