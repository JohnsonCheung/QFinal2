Attribute VB_Name = "MIde___Fun"
Option Explicit
Function CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Function
Function RmvTyChr$(A)
If IsTyChr(FstChr(A)) Then RmvTyChr = RmvFstChr(A) Else RmvTyChr = A
End Function

Function TmpPj() As VBProject
Set TmpPj = FxaPj(TmpFxa)
End Function

Function HasBar(Nm$)
HasBar = CurVbeHasBar(Nm)
End Function


Function AyWhCdLin(A) As String()
Dim L
For Each L In AyNz(A)
    If IsCdLin(L) Then PushI AyWhCdLin, L
Next
End Function


Function CdWinAy() As VBIDE.Window()
CdWinAy = WinTyWinAy(vbext_wt_CodeWindow)
End Function

Sub RenMd(NewNm$)
CurMd.Name = NewNm
End Sub

Sub SetActPj(PjNm$)
Set CurVbe.ActiveVBProject = Pj(PjNm)
End Sub

Private Sub Z_FbMthNy()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
'    For Each Fb In AppFbAy
        PushAy O, FbMthNy(Fb)
'    Next
    Brw O
    Return
X_BrwOne:
'    Brw FbMthNy(AppFbAy()(0))
    Return
End Sub


Sub Ens_Vbe_ZZDashPubMthAsPrivate()
VbeEnsZZDashPubMthAsPrivate CurVbe
End Sub

Sub EnsPrpOnEr()
MdEnsPrpOnEr CurMd
End Sub

Sub EnsSchmPrpOnEr()
MdEnsPrpOnEr Md("Schm")
MdEnsPrpOnEr Md("SchmT")
MdEnsPrpOnEr Md("SchmF")
End Sub

Function FbMthNy(A) As String()
FbMthNy = VbeMthNy(FbAcs(A).Vbe)
End Function
