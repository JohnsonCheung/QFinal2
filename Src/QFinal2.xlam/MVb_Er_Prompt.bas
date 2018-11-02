Attribute VB_Name = "MVb_Er_Prompt"
Option Explicit
Sub Done()
MsgBox "Done"
End Sub

Sub DtaEr()
MsgBox "DtaEr"
Stop
End Sub

Sub ErDta()
MsgBox "ErDta"
Stop
End Sub

Sub ErImposs()
Stop ' Impossible
End Sub

Sub ErNever()
MsgBox "Should never reach here"
Stop
End Sub

Sub ErPm()
MsgBox "Parameter Er"
Stop
End Sub

Sub ErRaise(Er$())
If AyBrwEr(Er) Then RaiseErr
End Sub

Function ErShow(Er$()) As String()
ErShow = SyShow("Er", Er)
End Function
