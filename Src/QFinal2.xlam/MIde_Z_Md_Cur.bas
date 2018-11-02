Attribute VB_Name = "MIde_Z_Md_Cur"
Option Explicit

Function CurMd() As CodeModule
Set CurMd = CurCdPne.CodeModule
End Function

Sub Z_CurMd()
Ass CurMd.Parent.Name = "Cur_d"
End Sub

Function CurMdDNm$()
CurMdDNm = MdDNm(CurMd)
End Function

Sub CurMdGenUBSz(TyNm$)
MdGenUBSz CurMd, TyNm
End Sub

Sub CurMdMovMth(MthPatn$, ToMd As CodeModule)
MdMthMov CurMd, MthPatn, ToMd
End Sub

Function CurMdMthNy() As String()
CurMdMthNy = MdMthNy(CurMd)
End Function

Function CurMdNm$()
CurMdNm = CurCmp.Name
End Function


Function CurMdWin() As VBIDE.Window
If IsNothing(CurVbe.ActiveCodePane) Then Exit Function
Set CurMdWin = CurVbe.ActiveCodePane.Window
End Function
