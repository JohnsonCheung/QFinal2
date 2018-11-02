Attribute VB_Name = "MVb_Run_CdLy"
Option Explicit
Const TarMdNm$ = "AAA_CdRun"
Private Sub ZAddMth(CdLy$(), MthNm$)
ZTarMd.AddFromString ZMthLines(CdLy, MthNm$)
End Sub
Private Function ZMthLines$(CdLy$(), MthNm$)
Dim Lines$, L1$, L2$
L1 = "Sub ZZZ_" & MthNm & "()"
Lines = JnCrLf(CdLy)
L2 = "End Sub"
ZMthLines = L1 & Lines & L2
End Function
Private Function ZTarMd() As CodeModule
Set ZTarMd = Application.Vbe.ActiveVBProject.VBComponents(TarMdNm).CodeModule
End Function

Sub RunCdLy(CdLy$())
RunCdLyMth CdLy, TmpNm
End Sub

Sub RunCdLyMth(CdLy$(), MthNm$)
ZAddMth CdLy, MthNm
Run "ZZZ_" & MthNm
End Sub

Private Function ZZCdLines$()
ZZCdLines = "MsgBox Now"
End Function
