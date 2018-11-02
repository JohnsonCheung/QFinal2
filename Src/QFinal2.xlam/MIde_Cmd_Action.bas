Attribute VB_Name = "MIde_Cmd_Action"
Option Explicit

Sub TileH()
CmdBtnOfTileH.Execute
End Sub

Sub TileV()
CmdBtnOfTileV.Execute
End Sub

Function TileVBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = WinPop.CommandBar.Controls(3)
If O.Caption <> "Tile &Vertically" Then Stop
Set TileVBtn = O
End Function
Sub ShwNxtStmt()
PopDbg.Visible = True
DoEvents
With NxtStmtBtn
'    If .Enabled Then .Execute
    .Execute
End With
End Sub
