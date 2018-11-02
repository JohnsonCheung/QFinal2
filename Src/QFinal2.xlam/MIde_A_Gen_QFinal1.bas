Attribute VB_Name = "MIde_A_Gen_QFinal1"
Option Explicit
Const FinalFxa$ = "C:\Users\User\Desktop\Vba-Lib-1\QFinal.xlam"

Sub GenQFinal1()
CpyCls
CpyMod
CpyAAAMod
PjSav Tar
End Sub

Private Sub CpyAAAMod()
PjMdDicApp Tar, AAAModDic
End Sub

Private Function Tar() As VBProject
Static X As VBProject
If IsNothing(X) Then Set X = FxaPj(TmpFxa(Fnn:="QFinal1"))
Set Tar = X
End Function

Private Function Src() As VBProject
Static X As VBProject
If IsNothing(X) Then Set X = Pj("QFinal")
Set Src = X
End Function

Private Sub CpyCls()
Dim M
For Each M In PjClsAy(Src)
    MdCpy CvMd(M), Tar
Next
End Sub

Private Function ModAy() As CodeModule()
ModAy = PjModAy(Src, WhNm(Exl:="AAAMod"))
End Function

Private Sub CpyMod()
Dim M
For Each M In ModAy
    MdCpy CvMd(M), Tar
Next
End Sub
