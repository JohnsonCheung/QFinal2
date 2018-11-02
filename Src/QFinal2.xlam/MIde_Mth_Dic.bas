Attribute VB_Name = "MIde_Mth_Dic"
Option Explicit
Sub Z()
Z_MdMthDic
Z_PjMthDic
Z_PjMthDic1
End Sub
Function MdMthDic(A As CodeModule) As Dictionary
Set MdMthDic = MdMthDDNmDic(A)
End Function

Private Sub ZZ_PjMthDic()
DicBrw PjMthDic(CurPj)
End Sub

Function PjMthDic(A As VBProject) As Dictionary
Dim I
Set PjMthDic = New Dictionary
For Each I In PjMdAy(A)
    PushDic PjMthDic, MdMthDic(CvMd(I))
Next
End Function

Private Sub ZZ_MdMthDic()
DicBrw MdMthDic(CurMd)
End Sub


Private Sub Z_MdMthDic()
DicBrw MdMthDic(CurMd)
End Sub
Private Sub Z_PjMthDic()
Dim A As Dictionary, V, K
Set A = PjMthDic(CurPj)
Ass IsSyDic(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub

Private Sub Z_PjMthDic1()
Dim A As Dictionary, V, K
Set A = PjMthDic(CurPj)
Ass IsSyDic(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub
