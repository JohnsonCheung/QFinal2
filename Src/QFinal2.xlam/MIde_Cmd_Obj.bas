Attribute VB_Name = "MIde_Cmd_Obj"
Option Explicit
Sub AssCompileBtn(PjNm$)
If CompileBtn.Caption <> "Compi&le " & PjNm Then Stop
End Sub

Sub BarClrAllCtl(A As CommandBar)
Dim I
For Each I In AyNz(BarCtlAy(A))
    CvCtl(I).Delete
Next
End Sub
Function IsBtn(A) As Boolean
IsBtn = TypeName(A) = "CommandButton"
End Function


Function BarCtlAy(A As CommandBar) As Control()
BarCtlAy = ItrAyInto(A.Controls, BarCtlAy)
End Function

Function BarCtlNy(A As CommandBar) As String()
End Function

Function PopDbg() As CommandBarPopup
Set PopDbg = MnuBar.Controls("Debug")
End Function

Function BarNy() As String()
BarNy = VbeBarNy(CurVbe)
End Function

Function CompileBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = PopDbg.CommandBar.Controls(1)
If Not HasPfx(O.Caption, "Compi&le") Then Stop
Set CompileBtn = O
End Function

Function WinPop() As CommandBarPopup
Set WinPop = MnuBar.Controls("Window")
End Function

Function IdeClrBtn() As Office.CommandBarButton
Set IdeClrBtn = ItrFstPrpEqV(PopEdt.Controls, "Caption", "C&lear")
End Function

Private Function PopEdt() As Office.CommandBarPopup
Set PopEdt = ItrFstPrpEqV(IdeMnuBar.Controls, "Caption", "&Edit")
End Function

Function IdeMnuBar() As Office.CommandBar
Set IdeMnuBar = CurVbe.CommandBars("Menu Bar")
End Function

Function IdeSelAllBtn() As Office.CommandBarButton
Set IdeSelAllBtn = ItrFstPrpEqV(PopEdt.Controls, "Caption", "Select &All")
End Function

Private Function MnuBar() As CommandBar
Set MnuBar = VbeMnuBar(CurVbe)
End Function

Function NxtStmtBtn() As CommandBarButton
Set NxtStmtBtn = PopDbg.Controls("Show Next Statement")
End Function

Function OCCwinVert() As Office.CommandBarButton
Set OCCwinVert = ItrFstPrpEqV(OCPwin.Controls, "Caption", "Tile &Vertically")
End Function

Function OCPwin() As Office.CommandBarPopup
Set OCPwin = ItrFstPrpEqV(IdeMnuBar.Controls, "Caption", "&Window")
End Function

Function SavBtn() As CommandBarButton
Dim I As CommandBarControl
For Each I In StdBar.Controls
    If HasPfx(I.Caption, "&Sav") Then Set SavBtn = I: Exit Function
Next
Stop
End Function

Sub ZZ_PopDbg()
Dim A
Set A = PopDbg
Stop
End Sub

Sub ZZ_MnuBar()
Dim A As CommandBar
Set A = MnuBar
Stop
End Sub

Function CmdBarCap_CmdPop(A As CommandBar, Cap$) As CommandBarPopup
Set CmdBarCap_CmdPop = ItrFstPrpEqV(A.Controls, "Caption", Cap)
End Function

Function CmdBarOfMnu() As CommandBar
Set CmdBarOfMnu = CurVbe.CommandBars("Menu Bar")
End Function

Sub Z_CmdBarOfMnu()
Debug.Print CmdBarOfMnu.Name
End Sub

Function CmdBtnOfTileH() As CommandBarButton
Set CmdBtnOfTileH = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Horizontally")
End Function

Function CmdBtnOfTileV() As CommandBarButton
Set CmdBtnOfTileV = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Vertically")
End Function

Function CmdPopCap_CmdBtn(A As CommandBarPopup, Cap$) As CommandBarButton
Set CmdPopCap_CmdBtn = ItrFstPrpEqV(A.Controls, "Caption", Cap)
End Function

Function CmdPopOfWin() As CommandBarPopup
Set CmdPopOfWin = CmdBarCap_CmdPop(CmdBarOfMnu, "&Window")
End Function

Sub Z_CmdPopOfWin()
Debug.Print CmdPopOfWin.Caption
End Sub

Sub CmdTurnOffTabStop(AcsCtl)
Dim A As Access.Control
Set A = AcsCtl
If Not HasPfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTgl(A): CvTgl(A).TabStop = False
End Select
End Sub
Function StdBar() As Office.CommandBar
Set StdBar = CurBars("Standard")
End Function

Function XlsBtn() As Office.CommandBarControl
Set XlsBtn = StdBar.Controls(1)
End Function
