Attribute VB_Name = "MIde_Cmd_MovMth"
Option Explicit
Const MovMthBarNm$ = "MovMth"
Const MovMthBtnNm$ = "MovMth"
Sub Z()
Z_MovMthBar
End Sub
Sub MovMthToMd(A$)
MsgBox A
End Sub

Private Sub EnsMovMthBtn()
EnsCmdBarBtn MovMthBarNm, MovMthBtnNm
MovMthBar.Visible = True
With MovMthBtn
    .Style = msoButtonIconAndCaption
    .OnAction = "MovMthToMd"
End With
End Sub
Sub EnsMovMthBar()
EnsCmdBar MovMthBarNm
EnsMovMthBtn
End Sub
Function CmdBarNy() As String()
CmdBarNy = ItrNy(CurBars)
End Function
Private Sub Z_MovMthBar()
MsgBox MovMthBar.Name
End Sub
Function CurBars() As Office.CommandBars
Set CurBars = CurVbe.CommandBars
End Function
Function CurBarsHas(A) As Boolean
CurBarsHas = ItrHasNm(CurBars, A)
End Function
Function CmdBar(A) As Office.CommandBar
Set CmdBar = CurBars(A)
End Function
Sub RmvCmdBar(A)
If CurBarsHas(A) Then CmdBar(A).Delete
End Sub
Function CvCmdBtn(A) As Office.CommandBarButton
Set CvCmdBtn = A
End Function
Function CmdBarHasBtn(A As Office.CommandBar, BtnCaption)
Dim C As Office.CommandBarControl
For Each C In A.Controls
    If C.Type = msoControlButton Then
        If CvCmdBtn(C).Caption = BtnCaption Then CmdBarHasBtn = True: Exit Function
    End If
Next
End Function
Sub EnsCmdBarBtn(CmdBarNm, BtnCaption)
EnsCmdBar MovMthBarNm
If CmdBarHasBtn(CmdBar(CmdBarNm), BtnCaption) Then Exit Sub
CmdBar(CmdBarNm).Controls.Add(msoControlButton).Caption = BtnCaption
End Sub
Sub EnsCmdBar(A$)
If CurBarsHas(A) Then Exit Sub
AddCmdBar A
End Sub
Sub AddCmdBar(A)
CurBars.Add A
End Sub
Function MovMthBar() As Office.CommandBar
Set MovMthBar = CurBars(MovMthBarNm)
End Function
Function MovMthBtn() As Office.CommandBarControl
Set MovMthBtn = MovMthBar.Controls(MovMthBtnNm)
End Function
