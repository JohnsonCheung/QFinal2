Attribute VB_Name = "MIde_Z_Win"
Option Explicit
Function LclWin() As VBIDE.Window
Set LclWin = WinTyWin(vbext_wt_Locals)
End Function
Function CurWin() As VBIDE.Window
Set CurWin = CurVbe.ActiveWindow
End Function

Function CurCdWin() As VBIDE.Window
Set CurCdWin = CurCdPne.Window
End Function

Function ImmWin() As VBIDE.Window
Set ImmWin = WinTyWin(vbext_wt_Immediate)
End Function

Function BrwObjWin() As VBIDE.Window
Set BrwObjWin = WinTyWin(vbext_wt_Browser)
End Function

Function VisWinCnt&()
VisWinCnt = ItrCntTruePrp(CurVbe.Windows, "Visible")
End Function

Function ApWinAy(ParamArray WinAp()) As VBIDE.Window()
Dim Av(): Av = WinAp
Dim I
For Each I In Av
    PushObj ApWinAy, I
Next
End Function

Function WinNy() As String()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    Debug.Print W.Caption, W.Type
    PushI WinNy, W.Caption
Next
End Function

Sub ClsWinExl(A As vbext_WindowType)
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    If W.Type <> A Then
        W.Close
    End If
Next
End Sub

Sub ClsWinExlImm()
ClsWinExl VBIDE.vbext_wt_Immediate
End Sub
Function CvWinAy(A) As VBIDE.Window()
CvWinAy = A
End Function
Function CvWin(A) As VBIDE.Window
Set CvWin = A
End Function

Sub AlignWinV()
OCCwinVert.Execute
End Sub

Sub WinClr(A As VBIDE.Window)
DoEvents
IdeSelAllBtn.Execute
DoEvents
SendKeys " "
'IdeClrBtn.Execute
End Sub

Sub WinCls(A As VBIDE.Window)
If IsNothing(A) Then Exit Sub
If WinTy(A) = -1 Then Exit Sub
If A.Visible Then A.Close
End Sub
Function EmpWinAy() As VBIDE.Window()
End Function

Sub ClsAllWinExl(ParamArray ExlWinAp())
Dim W As VBIDE.Window, Exl()
Exl = ExlWinAp
For Each W In CurVbe.Windows
    If Not OyHas(Exl, W) Then
        W.Visible = False
    End If
Next
Dim I
For Each I In Exl
    CvWin(I).Visible = True
Next
End Sub

Function MdWin(A As CodeModule) As VBIDE.Window
Set MdWin = A.CodePane.Window
End Function

Sub ClsCdWinExl(ExlMdNm$)
ClsAllWinExl MdWin(Md(ExlMdNm))
End Sub

Function WinCnt&()
WinCnt = Application.Vbe.Windows.Count
End Function

Function WinMdNm$(A As VBIDE.Window)
WinMdNm = TakBet(A.Caption, " - ", " (Code)")
End Function


Function WinTy(A As VBIDE.Window) As VBIDE.vbext_WindowType
On Error GoTo X
WinTy = A.Type
Exit Function
X: WinTy = -1
End Function

Function WinTyWin(A As vbext_WindowType) As VBIDE.Window
Set WinTyWin = ItrFstPrpEqV(CurVbe.Windows, "Type", A)
End Function

Function WinTyWinAy(T As vbext_WindowType) As VBIDE.Window()
WinTyWinAy = ItrWhPrpEqV(CurVbe.Windows, "Type", T)
End Function

Sub ClrImmWin()
With ImmWin
    .SetFocus
    .Visible = True
End With
DoEvents
SndKeys "^{HOME}^+{END} "
End Sub
Function WinVis(A As VBIDE.Window)
A.Visible = True
'A.WindowState
End Function
Sub ClsImmWin()
DoEvents
ImmWin.Visible = False
End Sub

Sub ShwDbg()
ClsAllWinExl ImmWin, LclWin, CurWin
Exit Sub
DoEvents
AlignWinV
Stop
ClrImmWin
End Sub

Sub ClsAllWin()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    W.Close
Next
End Sub
