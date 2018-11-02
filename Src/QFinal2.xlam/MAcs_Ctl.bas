Attribute VB_Name = "MAcs_Ctl"
Option Explicit

Sub TxtbSelPth(A As Access.TextBox)
Dim R$
R = PthSel(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub


Sub FrmSetCmdNotTabStop(A As Access.Form)
ItrDo A.Controls, "CmdTurnOffTabStop"
End Sub

Function CvCtl(A) As Access.Control
Set CvCtl = A
End Function

Function CvBtn(A) As Access.CommandButton
Set CvBtn = A
End Function


Function CvTgl(A) As Access.ToggleButton
Set CvTgl = A
End Function
